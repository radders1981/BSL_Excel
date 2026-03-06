"""
FastAPI application exposing the boring_semantic_layer models over HTTP.
All responses are JSON. CORS is open for local use.
"""
from __future__ import annotations

import ibis
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from app.loader import load_models

app = FastAPI(title="BSL Excel API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

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
        "dimensions": list(sm.dimensions.keys()),
        "measures": list(sm.measures.keys()),
    }


@app.post("/query")
def run_query(req: QueryRequest):
    if req.model not in MODELS:
        raise HTTPException(status_code=404, detail=f"Model '{req.model}' not found.")

    sm = MODELS[req.model]

    # Validate requested fields
    for d in req.dimensions:
        if d not in sm.dimensions:
            raise HTTPException(status_code=400, detail=f"Unknown dimension: '{d}'")
    for m in req.measures:
        if m not in sm.measures:
            raise HTTPException(status_code=400, detail=f"Unknown measure: '{m}'")

    # Build ibis filter expressions
    ibis_filters = []
    for f in req.filters:
        if f.dimension not in sm.dimensions:
            raise HTTPException(status_code=400, detail=f"Unknown filter field: '{f.dimension}'")
        if f.op not in FILTER_OPS:
            raise HTTPException(status_code=400, detail=f"Unknown operator: '{f.op}'. Use one of {sorted(FILTER_OPS)}")

        col_expr_fn = sm.dimensions[f.dimension]
        tbl = sm.table

        def make_filter(col_fn, op, val, tbl=tbl):
            col = col_fn(tbl)
            if op == "eq":
                return lambda t: col_fn(t) == _cast(col, val)
            elif op == "neq":
                return lambda t: col_fn(t) != _cast(col, val)
            elif op == "gt":
                return lambda t: col_fn(t) > _cast(col, val)
            elif op == "gte":
                return lambda t: col_fn(t) >= _cast(col, val)
            elif op == "lt":
                return lambda t: col_fn(t) < _cast(col, val)
            elif op == "lte":
                return lambda t: col_fn(t) <= _cast(col, val)
            elif op == "contains":
                return lambda t: col_fn(t).contains(val)

        ibis_filters.append(make_filter(col_expr_fn, f.op, f.value))

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
