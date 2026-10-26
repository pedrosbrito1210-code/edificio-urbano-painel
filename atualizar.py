"""Atualiza o painel publicado: baixa a versão mais nova do código, gera o
painel a partir da planilha e publica no Netlify se algo mudou.

Configuração fica em config_local.json (nunca vai para o GitHub):
{
  "excel": "C:/Users/pedro/OneDrive/.../Prestação de contas.xlsx",
  "titulo": "Nome do prédio",
  "netlify_site_id": "...",
  "netlify_token": "..."
}
"""
from __future__ import annotations

import hashlib
import io
import json
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

AQUI = Path(__file__).resolve().parent
CONFIG = AQUI / "config_local.json"
SITE = AQUI / "site"
HASH = AQUI / ".ultimo_publicado"


def carregar_config() -> dict:
    if not CONFIG.exists():
        sys.exit(f"ERRO: crie o arquivo {CONFIG.name} (veja config_local.exemplo.json).")
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    for chave in ("excel", "netlify_site_id", "netlify_token"):
        if not cfg.get(chave):
            sys.exit(f"ERRO: '{chave}' vazio em {CONFIG.name}.")
    return cfg


def baixar_codigo_novo() -> None:
    """git pull, se a pasta for um clone do repositório. Só lê do GitHub."""
    if not (AQUI / ".git").exists() or not shutil.which("git"):
        return
    r = subprocess.run(["git", "-C", str(AQUI), "pull", "--ff-only", "--quiet"], capture_output=True, text=True)
    if r.returncode != 0:
        print("Aviso: não consegui baixar a versão mais nova do código, sigo com a atual.")
        print("  ", r.stderr.strip()[:300])


def gerar(excel: str, titulo: str) -> None:
    r = subprocess.run([sys.executable, str(AQUI / "gerar_painel.py"), "--excel", excel,
                        "--saida", str(SITE / "index.html"), "--titulo", titulo])
    if r.returncode != 0:
        sys.exit(1)


def zip_do_site() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for arq in SITE.rglob("*"):
            if arq.is_file():
                z.write(arq, arq.relative_to(SITE).as_posix())
    return buf.getvalue()


def publicar(cfg: dict, conteudo: bytes) -> None:
    url = f"https://api.netlify.com/api/v1/sites/{cfg['netlify_site_id']}/deploys"
    req = urllib.request.Request(url, data=conteudo, method="POST", headers={
        "Authorization": f"Bearer {cfg['netlify_token']}",
        "Content-Type": "application/zip",
    })
    with urllib.request.urlopen(req, timeout=120) as resp:
        info = json.loads(resp.read())
    print(f"Publicado no Netlify: {info.get('ssl_url') or info.get('url')}  (deploy {info.get('id')})")


def main() -> int:
    cfg = carregar_config()
    baixar_codigo_novo()
    gerar(cfg["excel"], cfg.get("titulo", "Edifício"))

    conteudo = (SITE / "index.html").read_bytes()
    # ignora a linha de data para não republicar só porque a planilha foi aberta e salva sem mudanças
    assinatura = hashlib.sha256(b"\n".join(l for l in conteudo.splitlines() if b"geradoEm=" not in l)).hexdigest()
    if HASH.exists() and HASH.read_text().strip() == assinatura and "--forcar" not in sys.argv:
        print("Nada mudou desde a última publicação. Site mantido como está.")
        return 0

    try:
        publicar(cfg, zip_do_site())
    except Exception as e:  # noqa: BLE001
        print(f"ERRO ao publicar no Netlify: {e}")
        return 1
    HASH.write_text(assinatura)
    return 0


if __name__ == "__main__":
    sys.exit(main())
