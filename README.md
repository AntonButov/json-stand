# json-stand

JSON (pdf-stand shape) → Excel. CI converts each corpus file and runs **excel-stand** check (LibreOffice).

## Quick start

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
pip install "excel-stand @ git+https://github.com/AntonButov/excel-stand.git"
json-stand convert tests/corpus/pass/01_simple.json --out out.xlsx
```

## Deploy

**80.87.196.33:8132** only. See [docs/deploy.md](docs/deploy.md).

## License

MIT
