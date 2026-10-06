"""
Script de automação para compilação dos 4 executáveis usando PyInstaller.
"""

import sys
if sys.stdout:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr:
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import subprocess
import shutil
from pathlib import Path

diretorio_raiz = Path(__file__).resolve().parent
python_exe = sys.executable


import argparse


def fontes_foram_alterados(script_nome: str, exe_alvo: Path) -> bool:
    """Verifica se o script ou qualquer arquivo em core/ é mais recente que o executável compilado."""
    if not exe_alvo.exists():
        return True
    mtime_exe = exe_alvo.stat().st_mtime
    script_path = diretorio_raiz / script_nome
    if script_path.exists() and script_path.stat().st_mtime > mtime_exe:
        return True
    pasta_core = diretorio_raiz / "core"
    if pasta_core.exists():
        for f in pasta_core.glob("*.py"):
            if f.stat().st_mtime > mtime_exe:
                return True
    return False


def compilar(script_nome: str, nome_exe: str, add_data: list = None, forcar: bool = False):
    pasta_dist = diretorio_raiz / "dist"
    exe_alvo = pasta_dist / f"{nome_exe}.exe"

    if not forcar and exe_alvo.exists() and exe_alvo.stat().st_size > 1000000:
        if not fontes_foram_alterados(script_nome, exe_alvo):
            print(f"\n[INFO] {nome_exe}.exe já está atualizado em dist/ ({exe_alvo.stat().st_size // 1048576} MB). Pulando recompilação.")
            return

    print("\n" + "=" * 60)
    print(f"COMPILANDO: {nome_exe} a partir de {script_nome}...")
    print("=" * 60)

    cmd = [
        python_exe,
        "-m",
        "PyInstaller",
        "--onefile",
        "--console",
        "--clean",
        "--noconfirm",
        f"--name={nome_exe}",
    ]

    version_file = diretorio_raiz / "version_info.txt"
    if version_file.exists():
        cmd.append(f"--version-file={version_file}")

    cmd.append(str(diretorio_raiz / script_nome))

    if add_data:
        for item in add_data:
            cmd.append(f"--add-data={item}")

    resultado = subprocess.run(cmd, cwd=str(diretorio_raiz))
    if resultado.returncode != 0:
        print(f"[ERRO] Falha ao compilar {nome_exe}")
        sys.exit(resultado.returncode)
    print(f"[OK] Compilação concluída: {nome_exe}")


def main():
    parser = argparse.ArgumentParser(description="Automação de compilação dos executáveis do Gerador de Certificados ANM.")
    parser.add_argument("--relatorio", "-r", action="store_true", help="Recompilar apenas o gerar_relatorio.exe e atualizar criar_curso.exe.")
    parser.add_argument("--force", "-f", action="store_true", help="Forçar recompilação ignorando executáveis existentes.")
    args = parser.parse_args()

    pasta_dist = diretorio_raiz / "dist"
    pasta_bin = diretorio_raiz / "_bin"
    pasta_bin.mkdir(parents=True, exist_ok=True)

    if args.relatorio:
        # 1. Compila apenas gerar_relatorio.exe
        compilar("gerar_relatorio.py", "gerar_relatorio", forcar=True)
        origem = pasta_dist / "gerar_relatorio.exe"
        if origem.exists():
            shutil.copy2(origem, pasta_bin / "gerar_relatorio.exe")
            print(f"Copiado para _bin: gerar_relatorio.exe")

        # 2. Atualiza criar_curso.exe com o novo binário embutido
        compilar(
            "criar_curso.py",
            "criar_curso",
            add_data=["_bin;_bin", "templates;templates"],
            forcar=True,
        )
        origem_criar = pasta_dist / "criar_curso.exe"
        if origem_criar.exists():
            shutil.copy2(origem_criar, diretorio_raiz / "criar_curso.exe")
            print(f"Copiado para a raiz: criar_curso.exe")

        print("\n" + "=" * 60)
        print("[OK] gerar_relatorio.exe e criar_curso.exe RECOMPILADOS COM SUCESSO!")
        print("=" * 60)
        return

    # 1. Compila os executáveis das etapas (detectando fontes modificados ou se forçado)
    compilar("gerar_relatorio.py", "gerar_relatorio", forcar=args.force)
    compilar("gerar_certificados.py", "gerar_certificados", forcar=args.force)
    compilar("enviar_emails.py", "enviar_emails", forcar=args.force)

    # 2. Copia os executáveis das etapas para _bin/
    for exe_name in ["gerar_relatorio.exe", "gerar_certificados.exe", "enviar_emails.exe"]:
        origem = pasta_dist / exe_name
        if origem.exists():
            shutil.copy2(origem, pasta_bin / exe_name)
            print(f"Copiado para _bin: {exe_name}")

    # 3. Compila o executável principal criar_curso.exe com _bin e templates embutidos
    compilar(
        "criar_curso.py",
        "criar_curso",
        add_data=["_bin;_bin", "templates;templates"],
        forcar=True,
    )

    # 4. Copia criar_curso.exe para a raiz
    origem_criar = pasta_dist / "criar_curso.exe"
    if origem_criar.exists():
        shutil.copy2(origem_criar, diretorio_raiz / "criar_curso.exe")
        print(f"Copiado para a raiz: criar_curso.exe")

    print("\n" + "=" * 60)
    print("TODOS OS EXECUTÁVEIS FORAM COMPILADOS COM SUCESSO!")
    print("=" * 60)
    print(f"Executável raiz:      {diretorio_raiz / 'criar_curso.exe'}")
    print(f"Executáveis de apoio: {pasta_bin}")


if __name__ == "__main__":
    main()
