"""Safely inspect and download selected NASA PSI metadata from the public S3 bucket."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Iterator

BUCKET = "nasa-psi"
REGION = "us-west-2"
ENDPOINT = f"https://{BUCKET}.s3.{REGION}.amazonaws.com/"
TABLE_EXTENSIONS = {".csv", ".xlsx", ".xls", ".tsv"}
DOCUMENT_EXTENSIONS = {".pdf", ".doc", ".docx", ".ppt", ".pptx", ".txt", ".md"}
SKIP_EXTENSIONS = {".mp4", ".mov", ".avi", ".png", ".jpg", ".jpeg", ".tif", ".tiff"}
MAX_DOWNLOAD_BYTES = 250 * 1024 * 1024
NAMESPACE = {"s3": "http://s3.amazonaws.com/doc/2006-03-01/"}


def list_objects(prefix: str = "", delimiter: str | None = None) -> Iterator[dict[str, Any]]:
    """Yield public S3 objects without downloading their contents."""
    token: str | None = None
    while True:
        params = {"list-type": "2", "prefix": prefix, "max-keys": "1000"}
        if delimiter:
            params["delimiter"] = delimiter
        if token:
            params["continuation-token"] = token
        url = ENDPOINT + "?" + urllib.parse.urlencode(params)
        try:
            with urllib.request.urlopen(url, timeout=30) as response:
                root = ET.fromstring(response.read())
        except (OSError, ET.ParseError) as error:
            raise RuntimeError(f"Falha ao consultar o bucket público: {error}") from error

        for item in root.findall("s3:Contents", NAMESPACE):
            key = item.findtext("s3:Key", default="", namespaces=NAMESPACE)
            size = int(item.findtext("s3:Size", default="0", namespaces=NAMESPACE))
            yield {"key": key, "size": size}
        for item in root.findall("s3:CommonPrefixes", NAMESPACE):
            key = item.findtext("s3:Prefix", default="", namespaces=NAMESPACE)
            yield {"prefix": key}

        truncated = root.findtext("s3:IsTruncated", default="false", namespaces=NAMESPACE)
        token = root.findtext("s3:NextContinuationToken", namespaces=NAMESPACE)
        if truncated != "true" or not token:
            return


def public_url(key: str) -> str:
    return ENDPOINT + urllib.parse.quote(key, safe="/")


def human_size(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if value < 1024 or unit == "TB":
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def command_list(args: argparse.Namespace) -> None:
    """List prefixes or files, with no local writes."""
    for item in list_objects(args.prefix, "/" if args.top else None):
        if "prefix" in item:
            print(f"[PASTA] {item['prefix']}")
        else:
            print(f"{human_size(item['size']):>10}  {item['key']}")


def selected_key(key: str, include: re.Pattern[str] | None) -> bool:
    lower = key.lower()
    extension = Path(lower).suffix
    if extension in SKIP_EXTENSIONS:
        return False
    if extension not in TABLE_EXTENSIONS | DOCUMENT_EXTENSIONS:
        return False
    return include is None or include.search(key) is not None


def command_download(args: argparse.Namespace) -> None:
    include = re.compile(args.match, re.IGNORECASE) if args.match else None
    objects = [
        item
        for item in list_objects(args.prefix, None if args.recursive else "/")
        if "key" in item and selected_key(item["key"], include)
    ]
    if not objects:
        raise SystemExit("Nenhum CSV/XLSX/documento encontrado com os filtros informados.")
    if len(objects) > args.max_files:
        raise SystemExit(
            f"{len(objects)} arquivos foram encontrados; limite atual é {args.max_files}. "
            "Use --max-files explicitamente se isso for intencional."
        )

    destination = Path(args.output)
    for item in objects:
        size = int(item["size"])
        if size > args.max_bytes:
            print(f"IGNORADO (tamanho): {item['key']} ({human_size(size)})", file=sys.stderr)
            continue
        key = item["key"]
        target = destination / Path(key)
        target.parent.mkdir(parents=True, exist_ok=True)
        print(f"Baixando {key} -> {target} ({human_size(size)})")
        try:
            with urllib.request.urlopen(public_url(key), timeout=60) as response:
                with target.open("wb") as output:
                    remaining = size
                    while remaining:
                        chunk = response.read(min(1024 * 1024, remaining))
                        if not chunk:
                            raise RuntimeError(
                                f"download incompleto para {key}: "
                                f"{size - remaining} de {size} bytes"
                            )
                        output.write(chunk)
                        remaining -= len(chunk)
        except OSError as error:
            raise RuntimeError(f"Falha ao baixar {key}: {error}") from error


def scalar(value: Any) -> Any:
    if value is None or value == "":
        return None
    if hasattr(value, "item"):
        value = value.item()
    return value if isinstance(value, (str, int, float, bool)) else str(value)


def analyze_rows(rows: list[dict[str, Any]]) -> dict[str, Any]:
    columns = sorted({str(column) for row in rows for column in row})
    varying: dict[str, Any] = {}
    constant: dict[str, Any] = {}
    for column in columns:
        values = [scalar(row.get(column)) for row in rows]
        values = [value for value in values if value is not None]
        distinct = list(dict.fromkeys(json.dumps(value, sort_keys=True) for value in values))
        decoded = [json.loads(value) for value in distinct]
        if len(decoded) == 1:
            constant[column] = decoded[0]
        elif len(decoded) > 1:
            varying[column] = {"n_valores": len(decoded), "valores": decoded[:100]}
    return {"n_linhas": len(rows), "variou": varying, "constante": constant}


def read_table(path: Path) -> list[dict[str, Any]]:
    suffix = path.suffix.lower()
    if suffix in {".csv", ".tsv"}:
        delimiter = "\t" if suffix == ".tsv" else ","
        with path.open("r", encoding="utf-8-sig", newline="") as file:
            return list(csv.DictReader(file, delimiter=delimiter))
    if suffix in {".xlsx", ".xls"}:
        try:
            import pandas as pd
        except ImportError as error:
            raise RuntimeError(
                "Para XLS/XLSX, instale as dependências com: python -m pip install pandas openpyxl"
            ) from error
        frame = pd.read_excel(path)
        return frame.where(frame.notna(), None).to_dict(orient="records")
    raise ValueError(f"Formato de tabela não suportado: {path}")


def command_analyze(args: argparse.Namespace) -> None:
    root = Path(args.input)
    result: dict[str, Any] = {"source": "NASA Physical Sciences Informatics", "investigations": []}
    for path in sorted(root.rglob("*")):
        if path.is_file() and path.suffix.lower() in TABLE_EXTENSIONS:
            rows = read_table(path)
            result["investigations"].append(
                {
                    "id": path.stem,
                    "arquivo": str(path.relative_to(root)),
                    "fonte": public_url("/".join(path.relative_to(root).parts)),
                    "experimentos": analyze_rows(rows),
                }
            )
            print(f"Analisado: {path} ({len(rows)} linhas)")
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"JSON gerado: {output}")


def parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    listing = subparsers.add_parser("list", help="listar pastas ou arquivos")
    listing.add_argument("--prefix", default="", help="prefixo S3, por exemplo OASIS/")
    listing.add_argument("--top", action="store_true", help="mostrar somente o próximo nível de pastas")
    listing.set_defaults(function=command_list)

    download = subparsers.add_parser("download", help="baixar somente tabelas e documentos")
    download.add_argument("--prefix", required=True, help="prefixo escolhido, por exemplo OASIS/")
    download.add_argument("--output", default="data/raw", help="pasta local de destino")
    download.add_argument("--match", help="regex opcional para filtrar nomes de arquivos")
    download.add_argument(
        "--recursive",
        action="store_true",
        help="procurar em subpastas; use somente em um prefixo bem delimitado",
    )
    download.add_argument("--max-files", type=int, default=25)
    download.add_argument("--max-bytes", type=int, default=MAX_DOWNLOAD_BYTES)
    download.set_defaults(function=command_download)

    analyze = subparsers.add_parser("analyze", help="calcular colunas constantes e variáveis")
    analyze.add_argument("--input", default="data/raw")
    analyze.add_argument("--output", default="data/investigations.json")
    analyze.set_defaults(function=command_analyze)
    return parser


if __name__ == "__main__":
    try:
        arguments = parser().parse_args()
        arguments.function(arguments)
    except RuntimeError as error:
        raise SystemExit(str(error)) from error
