import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

path = sys.argv[1] if len(sys.argv) > 1 else "has_sections.json"
data = json.load(open(path, encoding="utf-8"))

print(f"{path}")
print(f"top level: {type(data).__name__}, {len(data):,} entries\n")

if isinstance(data, dict):
    keys = list(data)[:3]
    print(f"first keys: {keys}\n")
    first = data[keys[0]]
    print(f"value type: {type(first).__name__}")
    if isinstance(first, dict):
        print(f"value keys: {list(first)}")
        for k, v in first.items():
            if isinstance(v, list):
                print(f"\n  '{k}' is a list of {len(v)}")
                if v and isinstance(v[0], dict):
                    print(f"  item keys: {list(v[0])}")
                    for ik, iv in v[0].items():
                        shown = str(iv)[:70].replace("\n", " ")
                        print(f"    {ik:<12} {type(iv).__name__:<6} {shown}")
            else:
                print(f"  '{k}': {type(v).__name__} = {str(v)[:60]}")
    elif isinstance(first, list):
        print(f"value is a list of {len(first)}")
        if first and isinstance(first[0], dict):
            print(f"item keys: {list(first[0])}")
            for ik, iv in first[0].items():
                print(f"  {ik:<12} {type(iv).__name__:<6} "
                      f"{str(iv)[:70]}")

text_fields = set()


def walk(node, depth=0):
    if depth > 4:
        return
    if isinstance(node, dict):
        for k, v in node.items():
            if isinstance(v, str) and len(v) > 400:
                text_fields.add(k)
            walk(v, depth + 1)
    elif isinstance(node, list):
        for item in node[:3]:
            walk(item, depth + 1)


for key in list(data)[:40] if isinstance(data, dict) else []:
    walk(data[key])

print(f"\nfields holding long text: {sorted(text_fields) or 'NONE FOUND'}")
if not text_fields:
    print("-> this file does not store section text; the quotes have to come")
    print("   from the PDFs in has_ct_pdfs/ instead.")
