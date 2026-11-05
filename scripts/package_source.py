"""Package deliverable source plus prebuilt frontend, excluding runtime data."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

root = Path(__file__).resolve().parents[1]
excluded = {'.venv', '.venv312', 'node_modules', '__pycache__', '.pytest_cache', 'logs', 'data'}
target = root.parent / 'payroll-demo-source.zip'
with ZipFile(target, 'w', ZIP_DEFLATED) as archive:
    for path in root.rglob('*'):
        if path.is_file() and not excluded.intersection(path.relative_to(root).parts):
            archive.write(path, Path(root.name) / path.relative_to(root))
print(target)
