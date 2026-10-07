"""Prepare the repository launchers without reinstalling compatible dependencies."""
from __future__ import annotations

import os
import importlib
import site
import subprocess
import sys
import tomllib
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

ROOT = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parents[1]


def prepare(*extras: str) -> Path:
    if getattr(sys, 'frozen', False):
        os.environ.setdefault('FINANCIAL_SOURCES_DATA_ROOT', str(ROOT / 'data' / 'sources'))
        return ROOT
    try:
        if not 24 <= int(version('packaging').split('.')[0]) < 27:
            raise PackageNotFoundError('packaging')
    except PackageNotFoundError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'packaging>=24,<27'])
    # pip may create the user site after this interpreter has initialized.
    if site.ENABLE_USER_SITE:
        site.addsitedir(site.getusersitepackages())
    importlib.invalidate_caches()
    from packaging.requirements import Requirement
    project = tomllib.loads((ROOT / 'pyproject.toml').read_text(encoding='utf-8'))['project']
    requirements = list(project['dependencies'])
    for extra in extras:
        requirements.extend(project['optional-dependencies'][extra])
    missing = []
    for text in dict.fromkeys(requirements):
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
    if site.ENABLE_USER_SITE:
        site.addsitedir(site.getusersitepackages())
    importlib.invalidate_caches()
    sys.path.insert(0, str(ROOT / 'src'))
    return ROOT
