"""Run the current shared SBS scraper from the historical entry point."""
from __future__ import annotations

import os
import subprocess
import sys
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path


def ensure_dependencies(root: Path):
    import tomllib
    try:
        installed = version('packaging')
        major = int(installed.split('.')[0])
        if not 24 <= major < 27:
            raise PackageNotFoundError('packaging')
    except PackageNotFoundError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'packaging>=24,<27'])
    from packaging.requirements import Requirement
    requirements = tomllib.loads((root / 'pyproject.toml').read_text(encoding='utf-8'))['project']['dependencies']
    missing = []
    for text in requirements:
        requirement = Requirement(text)
        if requirement.marker and not requirement.marker.evaluate():
            continue
        try:
            installed = version(requirement.name)
            compatible = requirement.specifier.contains(installed)
        except PackageNotFoundError:
            installed, compatible = 'no instalado', False
        print(f'[deps] {requirement.name}: {installed} -> {"OK" if compatible else "ajustando"}')
        if not compatible:
            missing.append(text)
    if missing:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing])


def main() -> int:
    frozen = getattr(sys, 'frozen', False)
    root = Path(sys.executable).resolve().parent if frozen else Path(__file__).resolve().parents[2]
    if not frozen:
        ensure_dependencies(root)
        sys.path.insert(0, str(root / 'src'))
    os.environ.setdefault('FINANCIAL_SOURCES_DATA_ROOT', str(root / 'data' / 'sources'))
    from fuentes_financieras.cli_rates import main as run
    return run()


if __name__ == '__main__':
    raise SystemExit(main())
