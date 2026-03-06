"""
semantic_config.py — Edit this file to define your semantic models.

Each SemanticModel wraps a DuckDB table and declares:
  - dimensions: columns or expressions used to group/slice data
  - measures:   aggregations (sum, count, avg, etc.)

The MODELS dict at the bottom is required — main.py and the API import from it.
"""
import ibis
from boring_semantic_layer import SemanticModel

# Connect to DuckDB (change the path if you use a different file)
conn = ibis.duckdb.connect("data/tpch.duckdb")

# --- Orders model ---
orders = SemanticModel(
    table=conn.table("orders"),
    dimensions={
        "order_date": lambda t: t.o_orderdate,
        "status":     lambda t: t.o_orderstatus,
        "priority":   lambda t: t.o_orderpriority,
        "clerk":      lambda t: t.o_clerk,
    },
    measures={
        "total_price":  lambda t: t.o_totalprice.sum(),
        "order_count":  lambda t: t.count(),
        "avg_price":    lambda t: t.o_totalprice.mean(),
    },
)

# --- Line Item model ---
lineitem = SemanticModel(
    table=conn.table("lineitem"),
    dimensions={
        "ship_mode":    lambda t: t.l_shipmode,
        "return_flag":  lambda t: t.l_returnflag,
        "line_status":  lambda t: t.l_linestatus,
        "ship_date":    lambda t: t.l_shipdate,
    },
    measures={
        "quantity":         lambda t: t.l_quantity.sum(),
        "revenue":          lambda t: t.l_extendedprice.sum(),
        "avg_discount":     lambda t: t.l_discount.mean(),
        "line_count":       lambda t: t.count(),
    },
)

# REQUIRED: the API reads this dict to discover available models
MODELS = {
    "orders":   orders,
    "lineitem": lineitem,
}
