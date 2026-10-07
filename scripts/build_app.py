"""Build any supported SBS application in an isolated environment."""
from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGETS = {
    'passive-rates': ('scripts/sync_tasas_pasivas.py', 'TasasPasivasSBS', True),
    'average-exchange-rate': ('apps/sbs/tipo_cambio_promedio.py', 'TipoCambioSBS', False),
    'weighted-exchange-rate': ('apps/sbs/tipo_cambio_ponderado.py', 'TipoCambioPonderadoSBS', False),
}


def main(argv=None):
    parser = argparse.ArgumentParser(description='Empaquetador único de las aplicaciones SBS.')
    parser.add_argument('--app', required=True, choices=sorted(TARGETS))
    args = parser.parse_args(argv)
    script, name, console = TARGETS[args.app]
    source = ROOT / script
    if not source.is_file():
        raise FileNotFoundError(source)
    with tempfile.TemporaryDirectory(prefix='fuentes_app_build_') as temporary:
        directory = Path(temporary)
        env = directory / 'venv'
        subprocess.check_call([sys.executable, '-m', 'venv', str(env)])
        python = env / ('Scripts/python.exe' if sys.platform == 'win32' else 'bin/python')
        project = str(ROOT) if console else f'{ROOT}[desktop]'
        subprocess.check_call([str(python), '-m', 'pip', 'install', project, 'pyinstaller>=6,<7'])
        command = [str(python), '-m', 'PyInstaller', '--noconfirm', '--onefile',
                   '--console' if console else '--noconsole', '--name', name,
                   '--distpath', str(ROOT / 'dist'), '--workpath', str(directory / 'build'),
                   '--specpath', str(directory), '--paths', str(ROOT / 'scripts')]
        packages = ['fuentes_financieras', 'curl_cffi', 'pyarrow', 'pandas', 'numpy', 'bs4', 'lxml', 'openpyxl', 'certifi']
        if not console:
            packages.append('selenium')
        for package in packages:
            command.extend(['--collect-all', package])
        command.append(str(source))
        subprocess.check_call(command)
    print(f'Ejecutable generado en {ROOT / "dist"}.')


if __name__ == '__main__':
    main()
