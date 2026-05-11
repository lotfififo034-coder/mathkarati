# مذكرتي Pro v17 — Full Architectural Rebuild

## Architecture

```
mathkarati-v17/
├── app.py                 # Flask API (thin HTTP adapter only)
├── core/
│   ├── models.py          # Domain models + validation
│   └── themes.py          # All 12 color themes (immutable)
├── engine/
│   ├── primitives.py      # Low-level drawing (rect, text, gradient…)
│   ├── slides.py          # Slide builders (make_cover, make_results…)
│   └── pipeline.py        # Export orchestrator (single entry point)
├── public/
│   └── index.html         # Frontend (unchanged UX)
├── requirements.txt
├── build.sh
└── render.yaml
```

## What changed from v16

| v16 | v17 |
|-----|-----|
| 3 engines (canva, classic, node) | 1 unified Python engine |
| ~4500 lines across 3 generator files | ~1200 lines, clean separation |
| Threading + subprocess fragility | No threads, no subprocess |
| Temp files on disk | In-memory BytesIO only |
| Base64 transport (kept) | Base64 transport (kept — it works) |
| Fallback chaos | No fallbacks needed |
| Hard to debug | Each layer independently testable |

## Design principles

- **Deterministic**: same input → same output, every time
- **No side effects**: generation is pure (in, bytes out)
- **Fail fast**: validation before any I/O
- **Memory-safe**: pptx.save() → BytesIO, never disk
- **Stream-safe**: BytesIO is fully read before return
- **Thread-safe**: pipeline is stateless after font detection

## Deploy

```bash
bash build.sh
gunicorn app:app --bind 0.0.0.0:5000 --workers 2 --timeout 120
```
