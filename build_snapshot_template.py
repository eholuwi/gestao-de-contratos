#!/usr/bin/env python3
"""
build_snapshot_template.py
Gera assets/snapshot-template.js a partir dos arquivos-fonte do dashboard.

Por quê: o botão "Salvar" (assets/app.js) precisa montar um HTML standalone
(sem <script src> externos) para distribuir por e-mail/rede. Para isso ele
embute, em base64, o CSS, o Chart.js, o plugin de datalabels, regras.js,
app.js e o próprio HTML do shell (div#app) — ver buildSnapshotHtml() em
assets/app.js. Base64 evita que qualquer conteúdo (inclusive um comentário
de contrato com o texto "</script>") feche a tag de script prematuramente
quando o snapshot for aberto (base64 nunca contém "<" ou ">").

Rodar de novo sempre que theme.css, regras.js, app.js ou o shell de
Dashboard_Contratos.html mudarem — senão o snapshot gerado pelo botão
"Salvar" fica com uma versão desatualizada da aplicação (os DADOS sempre
saem atualizados; é só o "molde" do snapshot que é estático).

Uso:
    python build_snapshot_template.py
"""
import base64
import json
import pathlib
import re
import sys

BASE = pathlib.Path(__file__).parent
ASSETS = BASE / "assets"


def read_text(path):
    return path.read_text(encoding="utf-8")


def b64(text):
    return base64.b64encode(text.encode("utf-8")).decode("ascii")


def extract_body_html(html):
    m = re.search(r"<body>(.*?)<script", html, re.DOTALL)
    if not m:
        raise SystemExit("Não encontrei <body>...<script em Dashboard_Contratos.html")
    return m.group(1).strip()


def main():
    html = read_text(BASE / "Dashboard_Contratos.html")
    body_html = extract_body_html(html)

    payload = {
        "css": b64(read_text(ASSETS / "theme.css")),
        "bodyHtml": b64(body_html),
        "chartLib": b64(read_text(ASSETS / "chart.umd.min.js")),
        "datalabelsLib": b64(read_text(ASSETS / "chartjs-plugin-datalabels.min.js")),
        "regrasJs": b64(read_text(ASSETS / "regras.js")),
        "appJs": b64(read_text(ASSETS / "app.js")),
    }

    out = (
        "/* Gerado por build_snapshot_template.py — NÃO editar à mão.\n"
        "   Reexecute o script depois de mudar theme.css / regras.js / app.js /\n"
        "   o shell de Dashboard_Contratos.html. */\n"
        "window.SNAPSHOT_TEMPLATE = " + json.dumps(payload) + ";\n"
    )
    out_path = ASSETS / "snapshot-template.js"
    out_path.write_text(out, encoding="utf-8")
    print("OK ->", out_path, "(%.1f KB)" % (out_path.stat().st_size / 1024))


if __name__ == "__main__":
    sys.exit(main())
