#!/usr/bin/env python
"""
make_index_md.py
Gera index.md - índice legível em markdown para consulta rápida de endpoints.
Também gera search_index.json - índice flat para busca por palavra-chave.
"""
import json, os, re

SPLIT = "D:/Projetos/ss-hermes/spec/split"

with open(f"{SPLIT}/index.json", encoding="utf-8") as f:
    index = json.load(f)

# ── search_index.json: lista plana de todos os endpoints ─────────────────────
flat = []
for tag_entry in index["tags"]:
    tag = tag_entry["tag"]
    file = tag_entry["file"]
    for ep in tag_entry["endpoints"]:
        flat.append({
            "tag": tag,
            "file": file,
            "method": ep["method"],
            "path": ep["path"],
            "operationId": ep.get("operationId", ""),
            "summary": ep.get("summary", ""),
            "deprecated": ep.get("deprecated", False),
        })

with open(f"{SPLIT}/search_index.json", "w", encoding="utf-8") as f:
    json.dump(flat, f, ensure_ascii=False, indent=2)
print(f"search_index.json: {len(flat)} endpoints")

# ── index.md ──────────────────────────────────────────────────────────────────
lines = [
    f"# {index['title']}  v{index['version']}",
    "",
    f"**Tags:** {index['total_tags']}  |  **Paths:** {index['total_paths']}",
    "",
    "## Como usar",
    "- Consulte este índice para encontrar a tag/arquivo do endpoint desejado.",
    "- Leia `split/{file}` para obter paths + schemas completos.",
    "- Para busca por keyword: `split/search_index.json`.",
    "- Schemas globais: `split/components.json`.",
    "- Autenticação/info: `split/meta.json`.",
    "",
    "---",
    "",
    "## Endpoints por Tag",
    "",
]

for tag_entry in sorted(index["tags"], key=lambda x: x["tag"]):
    tag = tag_entry["tag"]
    file = tag_entry["file"]
    ep_count = tag_entry["endpoint_count"]
    lines.append(f"### {tag}  `({ep_count} endpoints)` → `{file}`")
    lines.append("")
    for ep in tag_entry["endpoints"]:
        dep = " ~~[deprecated]~~" if ep.get("deprecated") else ""
        summary = ep.get("summary") or ep.get("operationId") or ""
        lines.append(f"- `{ep['method']}` `{ep['path']}`{dep}  {summary}")
    lines.append("")

md = "\n".join(lines)
with open(f"{SPLIT}/index.md", "w", encoding="utf-8") as f:
    f.write(md)
print(f"index.md gerado ({len(md)//1024} KB)")
print("Concluido.")
