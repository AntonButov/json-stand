# Deploy json-stand

> Проверено: 2026-11-15

| | |
|--|--|
| Host | **80.87.196.33** only |
| Path | `/opt/json-stand` |
| Port | **8132** → container `8080` |
| Health | `http://80.87.196.33:8132/health` |

**Not on** `206.223.250.251`.

Secrets: `DEPLOY_SSH_PRIVATE_KEY`, `DEPLOY_USER` (`root`).

CI installs LibreOffice + [excel-stand](https://github.com/AntonButov/excel-stand) and runs convert→check on 20 JSON fixtures.
