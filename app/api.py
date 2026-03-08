"""
FastAPI application exposing the boring_semantic_layer models over HTTP.
All responses are JSON. CORS is open for local use.
"""
from __future__ import annotations

import traceback
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from typing import Literal

import pathlib

import ibis
from app.loader import load_models

import datetime, decimal

app = FastAPI(title="BSL Excel API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    traceback.print_exc()
    return JSONResponse(status_code=500, content={"error": str(exc)})


# Serve the task pane UI at /ui/
_docs_dir = pathlib.Path("docs")
if _docs_dir.exists():
    app.mount("/ui", StaticFiles(directory=str(_docs_dir), html=True), name="ui")

# Load models once at startup
MODELS = load_models()

# ---------------------------------------------------------------------------
# Request / response schemas
# ---------------------------------------------------------------------------

FILTER_OPS = {"eq", "neq", "gt", "gte", "lt", "lte", "contains"}

VALID_GRAINS = {"year", "quarter", "month", "date"}
# Map user-facing grain names to ibis truncate units
_GRAIN_TRUNCATE = {"year": "year", "quarter": "quarter", "month": "month", "date": "day"}


class FilterClause(BaseModel):
    dimension: str
    op: str
    value: str


class SortClause(BaseModel):
    field: str
    direction: Literal["asc", "desc"] = "asc"


class QueryRequest(BaseModel):
    model: str
    dimensions: list[str] = []
    measures: list[str] = []
    filters: list[FilterClause] = []
    sort_by: list[SortClause] = []
    grains: dict[str, Literal["year", "quarter", "month", "date"]] = {}
    limit: int = Field(default=1000, ge=1, le=100_000)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/debug/models")
def debug_models():
    """Inspect BSL SemanticModel schema methods — remove once working."""
    result = {}
    for name, sm in MODELS.items():
        dims = sm.get_dimensions()
        meas = sm.get_measures()
        result[name] = {
            "get_dimensions_type": type(dims).__name__,
            "get_dimensions_value": str(dims),
            "get_measures_type": type(meas).__name__,
            "get_measures_value": str(meas),
        }
    return result


@app.get("/models")
def list_models():
    return {"models": list(MODELS.keys())}


@app.get("/models/{model_name}/schema")
def model_schema(model_name: str):
    if model_name not in MODELS:
        raise HTTPException(status_code=404, detail=f"Model '{model_name}' not found.")
    sm = MODELS[model_name]
    tbl = sm.table
    all_dims = sm.get_dimensions()

    dim_info = []
    for name, dim_fn in all_dims.items():
        try:
            type_str = _ibis_type_to_str(dim_fn(tbl).type())
        except Exception:
            type_str = "string"
        dim_info.append({"name": name, "type": type_str})

    return {
        "model": model_name,
        "dimensions": dim_info,
        "measures": list(sm.get_measures().keys()),
    }


@app.post("/query")
def run_query(req: QueryRequest):
    if req.model not in MODELS:
        raise HTTPException(status_code=404, detail=f"Model '{req.model}' not found.")

    sm = MODELS[req.model]
    all_dims = sm.get_dimensions()
    all_meas = sm.get_measures()

    # Validate requested fields
    for d in req.dimensions:
        if d not in all_dims:
            raise HTTPException(status_code=400, detail=f"Unknown dimension: '{d}'")
    for m in req.measures:
        if m not in all_meas:
            raise HTTPException(status_code=400, detail=f"Unknown measure: '{m}'")

    # Validate grains
    for dim_name in req.grains:
        if dim_name not in req.dimensions:
            raise HTTPException(status_code=400, detail=f"Grain specified for '{dim_name}' which is not in the selected dimensions.")

    # Build ibis filter expressions
    ibis_filters = []
    for f in req.filters:
        if f.dimension not in all_dims:
            raise HTTPException(status_code=400, detail=f"Unknown filter field: '{f.dimension}'")
        if f.op not in FILTER_OPS:
            raise HTTPException(status_code=400, detail=f"Unknown operator: '{f.op}'. Use one of {sorted(FILTER_OPS)}")

        dim = all_dims[f.dimension]
        tbl = sm.table

        def make_filter(dim_obj, op, val, tbl=tbl):
            # Use dim_obj(tbl) to resolve both Deferred and callable exprs
            col = dim_obj(tbl)
            if op == "eq":
                return lambda t: dim_obj(t) == _cast(col, val)
            elif op == "neq":
                return lambda t: dim_obj(t) != _cast(col, val)
            elif op == "gt":
                return lambda t: dim_obj(t) > _cast(col, val)
            elif op == "gte":
                return lambda t: dim_obj(t) >= _cast(col, val)
            elif op == "lt":
                return lambda t: dim_obj(t) < _cast(col, val)
            elif op == "lte":
                return lambda t: dim_obj(t) <= _cast(col, val)
            elif op == "contains":
                return lambda t: dim_obj(t).contains(val)

        ibis_filters.append(make_filter(dim, f.op, f.value))

    # Validate sort fields — grained dimensions are output as "{dim}.{grain}"
    if req.sort_by:
        output_dim_names = {
            f"{d}.{req.grains[d]}" if d in req.grains else d
            for d in req.dimensions
        }
        valid_fields = output_dim_names | set(req.measures)
        for s in req.sort_by:
            if s.field not in valid_fields:
                raise HTTPException(
                    status_code=400,
                    detail=f"Sort field '{s.field}' is not in the query output."
                )

    try:
        if req.grains:
            result_df = _run_grained_query(sm, all_dims, all_meas, req, ibis_filters)
        else:
            order_by_param = [(s.field, s.direction) for s in req.sort_by] if req.sort_by else None
            query_expr = sm.query(
                dimensions=req.dimensions or None,
                measures=req.measures or None,
                filters=ibis_filters if ibis_filters else None,
                order_by=order_by_param,
                limit=req.limit,
            )
            result_df = query_expr.execute()
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))

    # Serialize to JSON-safe structure
    columns = list(result_df.columns)
    rows = []
    for row in result_df.itertuples(index=False):
        rows.append([_json_safe(v) for v in row])

    return {"columns": columns, "rows": rows}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _cast(col_expr, value: str):
    """Attempt to cast a string value to the column's dtype for comparisons."""
    dtype = col_expr.type()
    if dtype.is_numeric():
        try:
            return float(value) if "." in value else int(value)
        except ValueError:
            return value
    return value


def _json_safe(v):
    """Convert non-JSON-serialisable types (dates, Decimals) to strings."""
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.isoformat()
    if isinstance(v, decimal.Decimal):
        return float(v)
    return v


def _ibis_type_to_str(dtype) -> str:
    if dtype.is_date():
        return "date"
    if dtype.is_timestamp():
        return "timestamp"
    if dtype.is_integer():
        return "integer"
    if dtype.is_floating() or dtype.is_decimal():
        return "float"
    return "string"


def _run_grained_query(sm, all_dims, all_meas, req, ibis_filters):
    """Execute a query applying time grain truncations to the specified dimensions."""
    tbl = sm.table

    if ibis_filters:
        tbl = tbl.filter([f(tbl) for f in ibis_filters])

    dim_exprs = []
    for d in req.dimensions:
        col = all_dims[d](tbl)
        grain = req.grains.get(d)
        if grain:
            col = col.truncate(_GRAIN_TRUNCATE[grain])
            col_label = f"{d}.{grain}"
        else:
            col_label = d
        dim_exprs.append(col.name(col_label))

    agg_exprs = [all_meas[m](tbl).name(m) for m in req.measures]

    if dim_exprs and agg_exprs:
        query = tbl.group_by(dim_exprs).aggregate(agg_exprs)
    elif dim_exprs:
        query = tbl.select(dim_exprs).distinct()
    else:
        query = tbl.aggregate(agg_exprs)

    if req.sort_by:
        order_exprs = [
            ibis.desc(s.field) if s.direction == "desc" else s.field
            for s in req.sort_by
        ]
        query = query.order_by(order_exprs)

    import pandas as pd
    result_df = query.limit(req.limit).execute()

    # Ensure grained columns are serialized as "YYYY-MM-DD" (not full datetime strings)
    for dim_name, grain in req.grains.items():
        col_label = f"{dim_name}.{grain}"
        if col_label in result_df.columns:
            result_df[col_label] = pd.to_datetime(result_df[col_label]).dt.date

    return result_df
