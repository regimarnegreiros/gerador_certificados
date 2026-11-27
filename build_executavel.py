"""
Script de automação para gerar o executável desktop único (--onedir --noupx)
do Gerador de Certificados ANM usando gerador_certificados.spec.
"""

import argparse
from pathlib import Path
import shutil
import subprocess
import sys

if sys.stdout:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr:
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

DIRETORIO_RAIZ = Path(__file__).resolve().parent
SPEC_PATH = DIRETORIO_RAIZ / "gerador_certificados.spec"
PASTA_SAIDA_DIST = DIRETORIO_RAIZ / "dist" / "GeradorCertificadosANM"
EXE_DESTINO = PASTA_SAIDA_DIST / "GeradorCertificadosANM.exe"


def _calcular_tamanho_pasta_mb(pasta: Path) -> float:
    total = sum(f.stat().st_size for f in pasta.rglob("*") if f.is_file())
    return total / (1024 * 1024)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Compila o Gerador de Certificados ANM (GUI PySide6) em modo --onedir --noupx."
    )
    parser.add_argument(
        "--no-clean",
        action="store_true",
        help="Reutiliza cache de build/ quando possível sem limpar antes.",
    )
    args = parser.parse_args()

    if not SPEC_PATH.exists():
        print(f"[ERRO] Arquivo spec não encontrado: {SPEC_PATH}")
        return 1

    if not args.no_clean:
        for pasta in (DIRETORIO_RAIZ / "build" / "gerador_certificados", PASTA_SAIDA_DIST):
            if pasta.exists():
                shutil.rmtree(pasta, ignore_errors=True)

    print("=" * 64)
    print("COMPILANDO GERADOR DE CERTIFICADOS ANM (--onedir --noupx)...")
    print("=" * 64)

    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm"]
    if not args.no_clean:
        cmd.append("--clean")
    cmd.append(str(SPEC_PATH))

    res = subprocess.run(cmd, cwd=str(DIRETORIO_RAIZ))
    if res.returncode != 0:
        print(f"[ERRO] PyInstaller retornou código {res.returncode}.")
        return res.returncode

    if not EXE_DESTINO.exists():
        print(f"[ERRO] Executável esperado não foi encontrado em: {EXE_DESTINO}")
        return 1

    tam_exe_mb = EXE_DESTINO.stat().st_size / (1024 * 1024)
    tam_pasta_mb = _calcular_tamanho_pasta_mb(PASTA_SAIDA_DIST)
    qtd_arquivos = sum(1 for f in PASTA_SAIDA_DIST.rglob("*") if f.is_file())

    print("\n" + "=" * 64)
    print("[OK] COMPILAÇÃO CONCLUÍDA COM SUCESSO!")
    print("=" * 64)
    print(f"Pasta distribuível : {PASTA_SAIDA_DIST}")
    print(f"Executável         : {EXE_DESTINO} ({tam_exe_mb:.2f} MB)")
    print(f"Tamanho total      : {tam_pasta_mb:.1f} MB ({qtd_arquivos} arquivo(s))")
    print("\nPara distribuir aos usuários finais, compacte a pasta dist/GeradorCertificadosANM/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
