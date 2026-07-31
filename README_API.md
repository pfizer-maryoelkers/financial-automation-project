# Financial Automation — FastAPI Backend

The FastAPI layer exposes the pipeline as HTTP endpoints so any frontend
(Build Me / Next.js, Postman, curl, etc.) can call it.

---

## Install dependencies

```bash
pip install -r requirements.txt
```

---

## Start the server

```bash
uvicorn api:app --reload --port 8000
```

The interactive docs are then at **http://localhost:8000/docs**

---

## Endpoints

### `GET /health`
Liveness check.
```json
{ "status": "ok" }
```

---

### `POST /extract-cost-centers`
Upload a template file and get back the list of cost centers in it.
Use this to populate the cost center selector in the UI before running.

**Form data**
| Field | Type | Description |
|---|---|---|
| `template_file` | file | Template `.xlsx` |

**Response**
```json
{ "cost_centers": ["1234", "5678", "9012"] }
```

---

### `POST /run`
Run the full pipeline. Returns the completed output `.xlsx` as a file download.

**Form data**
| Field | Type | Description |
|---|---|---|
| `template_file` | file | Template `.xlsx` |
| `transactional_file` | file | Transactional detail `.xlsx` |
| `forecast_files` | file (repeat) | One or more forecast `.xlsx` files |
| `selected_cost_centers` | string (optional) | Comma-separated IDs e.g. `"1234,5678"` |

**Response** — binary `.xlsx` file download.

---

### `POST /run/summary`
Same inputs as `/run` but returns a **JSON summary** instead of a file.
Useful for Build Me to display run results (exception counts, logs) before
offering a download link.

**Response**
```json
{
  "status": "success",
  "output_filename": "template_output.xlsx",
  "execution_time_seconds": 12.4,
  "exceptions": {
    "counts": { "MISSING_PO": 3 },
    "total": 3,
    "percentages": { "MISSING_PO": 100.0 }
  },
  "logs": [
    { "level": "INFO", "timestamp": "14:03:01", "message": "Pipeline starting..." }
  ],
  "output_file_base64": "<base64 encoded xlsx>"
}
```

To decode the file in the Next.js frontend:
```typescript
const bytes = Buffer.from(data.output_file_base64, "base64");
const blob = new Blob([bytes], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
});
const url = URL.createObjectURL(blob);
```

---

## Calling from Next.js (Build Me)

### Extract cost centers
```typescript
// app/api/extract-cost-centers/route.ts
export async function POST(req: Request) {
  const formData = await req.formData();
  const res = await fetch("http://localhost:8000/extract-cost-centers", {
    method: "POST",
    body: formData,
  });
  return Response.json(await res.json());
}
```

### Run the pipeline
```typescript
// app/api/run/route.ts
export async function POST(req: Request) {
  const formData = await req.formData();
  const res = await fetch("http://localhost:8000/run/summary", {
    method: "POST",
    body: formData,
  });
  return Response.json(await res.json());
}
```

---

## Project structure

```
api.py              ← FastAPI app + all endpoints
api_backend.py      ← Pipeline orchestrator (no UI dependency)
api_config.py       ← Config loading (no Streamlit dependency)
src/                ← Core pipeline — unchanged
configs/            ← YAML configs — unchanged
```

`app.py` and `streamlit_backend.py` are the old Streamlit layer and are no
longer needed once you move to the API backend.
