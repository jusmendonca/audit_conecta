#!/usr/bin/env python
"""
split_spec.py
Divide o OpenAPI spec em arquivos menores e gera um índice de consulta.

Estrutura gerada em spec/split/:
  index.json          - Índice completo: tags, endpoints, métodos, summaries
  meta.json           - Info, openapi version, security schemes
  components.json     - Schemas completos de componentes
  tag_{TagName}.json  - Um arquivo por tag com paths + schemas referenciados
  _untagged.json      - Paths sem tag
"""

import json
import re
import os

SRC = "D:/Projetos/ss-hermes/spec/supersapiens-spec.json"
DEST = "D:/Projetos/ss-hermes/spec/split"

os.makedirs(DEST, exist_ok=True)

print("Carregando spec...")
with open(SRC, "r", encoding="utf-8") as f:
    spec = json.load(f)

info = spec.get("info", {})
paths = spec.get("paths", {})
components = spec.get("components", {})
schemas = components.get("schemas", {})
security_schemes = components.get("securitySchemes", {})
tags_meta = spec.get("tags", [])

# ── 1. meta.json ──────────────────────────────────────────────────────────────
meta = {
    "openapi": spec.get("openapi"),
    "info": info,
    "securitySchemes": security_schemes,
    "tags": tags_meta,
}
with open(f"{DEST}/meta.json", "w", encoding="utf-8") as f:
    json.dump(meta, f, ensure_ascii=False, indent=2)
print("ok meta.json")

# ── 2. components.json ────────────────────────────────────────────────────────
with open(f"{DEST}/components.json", "w", encoding="utf-8") as f:
    json.dump({"schemas": schemas}, f, ensure_ascii=False, indent=2)
print("ok components.json")

# ── 3. Agrupar paths por tag ──────────────────────────────────────────────────
METHODS = ("get", "post", "put", "patch", "delete", "head", "options")

tag_paths: dict[str, dict] = {}   # tag -> {path -> path_obj}
untagged_paths: dict = {}

for path, path_obj in paths.items():
    path_tags = set()
    for method, op in path_obj.items():
        if method in METHODS:
            for t in op.get("tags", []):
                path_tags.add(t)
    if not path_tags:
        untagged_paths[path] = path_obj
    for tag in path_tags:
        tag_paths.setdefault(tag, {})[path] = path_obj

# ── 4. Coletar refs de schema usados em um conjunto de paths ──────────────────
REF_RE = re.compile(r'"#/components/schemas/([^"]+)"')

def collect_refs(obj_str: str) -> set[str]:
    return set(REF_RE.findall(obj_str))

def resolve_refs_recursive(names: set[str], all_schemas: dict) -> dict:
    """Retorna todos schemas necessários incluindo dependências transitivas."""
    resolved = {}
    queue = list(names)
    while queue:
        name = queue.pop()
        if name in resolved or name not in all_schemas:
            continue
        schema = all_schemas[name]
        resolved[name] = schema
        # descobre refs dentro do schema
        sub_refs = collect_refs(json.dumps(schema))
        for r in sub_refs:
            if r not in resolved:
                queue.append(r)
    return resolved

# ── 5. Index structure ────────────────────────────────────────────────────────
index = {
    "title": info.get("title", ""),
    "version": info.get("version", ""),
    "total_tags": len(tag_paths),
    "total_paths": len(paths),
    "tags": []
}

safe = lambda name: re.sub(r'[^\w\-]', '_', name)

# ── 6. Gerar arquivo por tag ──────────────────────────────────────────────────
print(f"Processando {len(tag_paths)} tags...")
for tag, tag_path_obj in sorted(tag_paths.items()):
    paths_str = json.dumps(tag_path_obj)
    ref_names = collect_refs(paths_str)
    tag_schemas = resolve_refs_recursive(ref_names, schemas)

    # Endpoints para o índice
    endpoints = []
    for path, path_obj in tag_path_obj.items():
        for method, op in path_obj.items():
            if method in METHODS:
                endpoints.append({
                    "method": method.upper(),
                    "path": path,
                    "operationId": op.get("operationId", ""),
                    "summary": op.get("summary", ""),
                    "deprecated": op.get("deprecated", False),
                })

    # Arquivo da tag
    filename = f"tag_{safe(tag)}.json"
    tag_doc = {
        "tag": tag,
        "paths": tag_path_obj,
        "schemas": tag_schemas,
    }
    with open(f"{DEST}/{filename}", "w", encoding="utf-8") as f:
        json.dump(tag_doc, f, ensure_ascii=False, indent=2)

    index["tags"].append({
        "tag": tag,
        "file": filename,
        "endpoint_count": len(endpoints),
        "endpoints": endpoints,
    })

# ── 7. Paths sem tag ──────────────────────────────────────────────────────────
if untagged_paths:
    untagged_str = json.dumps(untagged_paths)
    ref_names = collect_refs(untagged_str)
    untagged_schemas = resolve_refs_recursive(ref_names, schemas)
    untagged_doc = {"tag": "_untagged", "paths": untagged_paths, "schemas": untagged_schemas}
    with open(f"{DEST}/_untagged.json", "w", encoding="utf-8") as f:
        json.dump(untagged_doc, f, ensure_ascii=False, indent=2)

    endpoints = []
    for path, path_obj in untagged_paths.items():
        for method, op in path_obj.items():
            if method in METHODS:
                endpoints.append({
                    "method": method.upper(),
                    "path": path,
                    "operationId": op.get("operationId", ""),
                    "summary": op.get("summary", ""),
                })
    index["tags"].append({
        "tag": "_untagged",
        "file": "_untagged.json",
        "endpoint_count": len(endpoints),
        "endpoints": endpoints,
    })
    print(f"✓ _untagged.json ({len(untagged_paths)} paths)")

# ── 8. Salvar índice ──────────────────────────────────────────────────────────
with open(f"{DEST}/index.json", "w", encoding="utf-8") as f:
    json.dump(index, f, ensure_ascii=False, indent=2)
print(f"✓ index.json ({len(index['tags'])} tags)")

# ── 9. Relatório ──────────────────────────────────────────────────────────────
files = os.listdir(DEST)
total_size = sum(os.path.getsize(f"{DEST}/{fn}") for fn in files)
print(f"\nTotal de arquivos gerados: {len(files)}")
print(f"Tamanho total: {total_size/1024/1024:.1f} MB")
print(f"Maior tag por endpoints:")
top = sorted(index["tags"], key=lambda x: -x["endpoint_count"])[:10]
for t in top:
    print(f"  {t['tag']}: {t['endpoint_count']} endpoints -> {t['file']}")
print("\nConcluído.")
