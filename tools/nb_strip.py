# Git clean filter: removes outputs and execution counts from a notebook read on stdin,
# so that data shown in cell outputs is never committed. Usage (set in .gitattributes):
#   git config filter.nbstrip.clean "python tools/nb_strip.py"
import json
import sys

nb = json.loads(sys.stdin.buffer.read().decode("utf-8"))
for cell in nb.get("cells", []):
    if cell.get("cell_type") == "code":
        cell["outputs"] = []
        cell["execution_count"] = None
sys.stdout.buffer.write((json.dumps(nb, ensure_ascii=False, indent=1) + "\n").encode("utf-8"))
