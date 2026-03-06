# BSL Excel

Query a [boring_semantic_layer](https://github.com/boringdata/boring-semantic-layer) DuckDB semantic model directly from Excel — no BI platform, no cloud subscription, no build toolchain.

```
DuckDB file
    └── semantic_config.py (your model definitions)
            └── FastAPI server (main.py)
                    └── Excel task pane (Office Add-in)
```

---

## Quick Start

### 1. Prerequisites

- Python 3.11+
- [uv](https://docs.astral.sh/uv/) (`pip install uv` or `winget install astral-sh.uv`)
- Microsoft Excel (Desktop or Online)

### 2. Clone & install

```bash
git clone https://github.com/YOUR_GITHUB_USERNAME/bsl-excel
cd bsl-excel
uv sync
```

### 3. Generate the demo database

```bash
uv run python scripts/generate_demo_db.py
```

This creates `data/tpch.duckdb` with TPC-H sample data (~12 MB).

### 4. Start the API server

```bash
uv run python main.py
```

The server starts at `http://localhost:8000`. Verify it's running:

```bash
curl http://localhost:8000/health
# {"status":"ok"}
```

### 5. Sideload the add-in into Excel

1. Open Excel
2. **Insert → Add-ins → Upload My Add-in**
3. Select `manifest.xml` from this project
4. The "BSL Excel" button appears in the **Home** ribbon

### 6. Query from Excel

1. Open the BSL Excel task pane
2. Enter `http://localhost:8000` as the server URL and click **Connect**
3. Select a model (e.g. `orders`)
4. Check the dimensions and measures you want
5. Optionally add a filter
6. Click a cell in your sheet, then click **Run Query → Sheet**

---

## Defining Your Own Semantic Model

Edit `semantic_config.py`. You need two things:

1. An ibis DuckDB connection pointing at your `.duckdb` file
2. One or more `SemanticModel` objects collected into a `MODELS` dict

```python
import ibis
from boring_semantic_layer import SemanticModel

conn = ibis.duckdb.connect("data/my_data.duckdb")

sales = SemanticModel(
    table=conn.table("sales"),
    dimensions={
        "region":    lambda t: t.region,
        "category":  lambda t: t.product_category,
        "sale_date": lambda t: t.sale_date,
    },
    measures={
        "revenue":       lambda t: t.amount.sum(),
        "order_count":   lambda t: t.count(),
        "avg_order":     lambda t: t.amount.mean(),
    },
)

MODELS = {"sales": sales}
```

Restart the server after editing (`Ctrl+C` then `uv run python main.py`).

---

## API Reference

| Method | Path | Description |
|--------|------|-------------|
| GET | `/health` | Server health check |
| GET | `/models` | List available model names |
| GET | `/models/{name}/schema` | Get dimensions + measures for a model |
| POST | `/query` | Execute a query, returns columns + rows |

### POST /query body

```json
{
  "model": "orders",
  "dimensions": ["order_date", "status"],
  "measures": ["total_price", "order_count"],
  "filters": [
    { "dimension": "status", "op": "eq", "value": "O" }
  ],
  "limit": 1000
}
```

Filter operators: `eq` `neq` `gt` `gte` `lt` `lte` `contains`

Interactive API docs available at `http://localhost:8000/docs` when the server is running.

---

## Publishing the Add-in UI to GitHub Pages

The `addin/` folder is a static web app that Excel loads. To host it publicly:

1. Push this repo to GitHub
2. **Settings → Pages → Source**: branch `main`, folder `/addin`
3. GitHub will provide a URL like `https://yourname.github.io/bsl-excel/addin/`
4. Edit `manifest.xml` — replace all `YOUR_GITHUB_USERNAME` and `YOUR_REPO_NAME` placeholders with your real values
5. Re-sideload the updated `manifest.xml` in Excel

Until you set up GitHub Pages you can also use a local server for development:

```bash
python -m http.server 3000 --directory addin
# Then set SourceLocation in manifest.xml to http://localhost:3000
```

---

## Project Structure

```
bsl-excel/
├── main.py                 # Start the API server
├── semantic_config.py      # Your semantic model definitions (edit this)
├── manifest.xml            # Excel add-in manifest (sideload into Excel)
├── pyproject.toml          # Python dependencies (managed by uv)
├── .python-version         # Python version pin
│
├── app/
│   ├── api.py              # FastAPI routes
│   └── loader.py           # Loads semantic_config.py dynamically
│
├── data/
│   └── tpch.duckdb         # Demo DuckDB database
│
├── addin/                  # Static Excel task pane (GitHub Pages)
│   ├── index.html
│   ├── taskpane.js
│   └── taskpane.css
│
└── scripts/
    └── generate_demo_db.py # One-time demo data generation
```

---

## Requirements

- `boring-semantic-layer >= 0.3.2`
- `fastapi >= 0.115`
- `uvicorn[standard] >= 0.30`
- `ibis-framework[duckdb] >= 9.0`

All managed automatically by `uv sync`.

---

## License

MIT
