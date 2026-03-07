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
from pydantic import BaseModel

import pathlib

from app.loader import load_models

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


class FilterClause(BaseModel):
    dimension: str
    op: str
    value: str


class QueryRequest(BaseModel):
    model: str
    dimensions: list[str] = []
    measures: list[str] = []
    filters: list[FilterClause] = []
    limit: int = 1000


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
    return {
        "model": model_name,
        "dimensions": list(sm.get_dimensions().keys()),
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

    try:
        result_df = sm.query(
            dimensions=req.dimensions or None,
            measures=req.measures or None,
            filters=ibis_filters if ibis_filters else None,
            limit=req.limit,
        ).execute()
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
    import datetime, decimal
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.isoformat()
    if isinstance(v, decimal.Decimal):
        return float(v)
    return v
