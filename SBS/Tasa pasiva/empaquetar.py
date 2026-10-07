"""Build a console executable for the current shared SBS scraper."""
from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path


def main():
    folder = Path(__file__).resolve().parent
    repo = folder.parents[1]
    with tempfile.TemporaryDirectory(prefix='sbs_rates_build_') as temporary:
        root = Path(temporary)
        env = root / 'venv'
        subprocess.check_call([sys.executable, '-m', 'venv', str(env)])
        python = env / ('Scripts/python.exe' if sys.platform == 'win32' else 'bin/python')
        subprocess.check_call([str(python), '-m', 'pip', 'install', str(repo), 'pyinstaller>=6,<7'])
        command = [str(python), '-m', 'PyInstaller', '--noconfirm', '--onefile', '--console',
                   '--name', 'TasasPasivasSBS', '--distpath', str(folder / 'dist'),
                   '--workpath', str(root / 'build'), '--specpath', str(root)]
        for package in ('fuentes_financieras', 'curl_cffi', 'pyarrow', 'pandas', 'numpy', 'bs4', 'lxml', 'openpyxl', 'certifi'):
            command.extend(['--collect-all', package])
        command.append(str(folder / 'script.py'))
        subprocess.check_call(command)
    print(f'Ejecutable generado en {folder / "dist"}. Use --help para las opciones.')


if __name__ == '__main__':
    main()
