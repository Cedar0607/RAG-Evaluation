Yes. The template is stored in the beta backend’s Docker volume at:

```text
/app/local_files/supplier_selection/template.pptx
```

Do not edit Docker’s `/var/lib/docker/volumes/...` directory directly. Use `docker compose cp` so the correct beta volume is targeted.

## 1. Transfer the PPTX to the Linux server

For example, place it temporarily at:

```text
/tmp/supplier-selection-template.pptx
```

Confirm it exists:

```bash
ls -lh /tmp/supplier-selection-template.pptx
```

Optionally verify that it is a valid PPTX/ZIP file:

```bash
unzip -t /tmp/supplier-selection-template.pptx
```

## 2. Copy it into the beta backend

Run from the beta deployment directory:

```bash
cd /opt/material-system-beta
```

Copy it to a temporary location inside the persistent volume:

```bash
docker compose --env-file .env cp \
  /tmp/supplier-selection-template.pptx \
  backend:/app/local_files/supplier_selection/template.uploading
```

## 3. Validate and activate the template

Run:

```bash
docker compose --env-file .env exec -T backend python - <<'PY'
from pathlib import Path
import os
import zipfile

directory = Path("/app/local_files/supplier_selection")
source = directory / "template.uploading"
destination = directory / "template.pptx"
name_file = directory / "template.name"

if not source.is_file():
    raise SystemExit("Uploaded template file was not found")

if not zipfile.is_zipfile(source):
    raise SystemExit("The supplied file is not a valid PPTX/ZIP file")

with zipfile.ZipFile(source) as archive:
    required = {"[Content_Types].xml", "ppt/presentation.xml"}
    missing = required.difference(archive.namelist())
    if missing:
        raise SystemExit(f"Invalid PPTX; missing: {sorted(missing)}")

os.replace(source, destination)
name_file.write_text(
    "supplier-selection-template.pptx",
    encoding="utf-8",
)

print(f"Template activated: {destination}")
PY
```

This performs an atomic replacement, reducing the risk of Auto-PPT reading a partially copied file.

## 4. Verify the stored template

```bash
docker compose --env-file .env exec -T backend \
  ls -lh /app/local_files/supplier_selection/
```

You should see:

```text
template.pptx
template.name
```

Verify the internal download endpoint:

```bash
docker compose --env-file .env exec -T auto_ppt python - <<'PY'
import urllib.request

url = "http://backend:8000/api/v1/supplier-selection/template"
response = urllib.request.urlopen(url, timeout=10)
data = response.read()

print("HTTP status:", response.status)
print("Template size:", len(data), "bytes")
print("PPTX signature valid:", data[:2] == b"PK")
PY
```

No container restart is required. The next Auto-PPT task will automatically use the new template.

Finally, remove the temporary host copy if it is no longer needed:

```bash
rm /tmp/supplier-selection-template.pptx
```

These commands affect only the `material-system-beta` template volume, provided they are run from `/opt/material-system-beta` with its beta Compose file.
