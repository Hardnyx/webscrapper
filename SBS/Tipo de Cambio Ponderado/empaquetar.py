import os
import sys
import time
import shutil
import tempfile
import subprocess
from pathlib import Path

# ==================== CONFIGURACIÓN ====================
PY_FILE  = "script.py"
APP_NAME = "TipoCambioPonderadoSBS"
# =======================================================

root = Path().resolve()
print("Carpeta del proyecto:", root)

src = root / PY_FILE
if not src.exists():
    raise FileNotFoundError(f"No se encontró {PY_FILE} en {root}")

tmp_root  = Path(tempfile.mkdtemp(prefix="tasas_build_"))
venv_dir  = tmp_root / "venv"
tmp_dist  = tmp_root / "dist"
tmp_build = tmp_root / "build"

print(f"Carpeta de trabajo temporal: {tmp_root}")

# ==================== PASO 1: Venv limpio en %TEMP% ====================
print("\n[1/4] Creando entorno virtual limpio en carpeta temporal...")
result = subprocess.run(
    [sys.executable, "-m", "venv", str(venv_dir)],
    capture_output=True, text=True
)
if result.returncode != 0:
    raise RuntimeError(f"No se pudo crear el venv: {result.stderr}")
print("  Venv creado.")

venv_python = (
    venv_dir / "Scripts" / "python.exe"
    if os.name == "nt"
    else venv_dir / "bin" / "python"
)

# ==================== PASO 2: Instalar paquetes en el venv ====================
print("\n[2/4] Instalando paquetes necesarios en el venv...")

packages = [
    "pyinstaller",
    "selenium",
    "requests",
    "beautifulsoup4",
    "lxml",
    "pandas",
    "numpy",
    "openpyxl",
]

result = subprocess.run(
    [str(venv_python), "-m", "pip", "install", "--quiet"] + packages,
    capture_output=True, text=True
)
if result.returncode != 0:
    shutil.rmtree(tmp_root, ignore_errors=True)
    raise RuntimeError(f"Error instalando paquetes: {result.stderr[:300]}")

for pkg in packages:
    print(f"  {pkg:25} OK")
print("  Instalación completada.")

# ==================== PASO 3: Preparar script fuente ====================
print("\n[3/4] Preparando script fuente...")

script_text = src.read_text(encoding="utf-8")

if "webdriver_manager" in script_text:
    print("  Revirtiendo parche de webdriver-manager (no necesario en Selenium 4.x)...")
    script_text = script_text.replace(
        "from selenium import webdriver\n"
        "from selenium.webdriver.chrome.service import Service\n"
        "from webdriver_manager.chrome import ChromeDriverManager",
        "from selenium import webdriver",
        1
    )
    script_text = script_text.replace(
        "driver = webdriver.Chrome(\n"
        "    service=Service(ChromeDriverManager().install())\n"
        ")",
        "driver = webdriver.Chrome()",
        1
    )
    print("  Parche revertido.")
else:
    print("  Script listo. Usará selenium-manager integrado de Selenium 4.x.")

build_src = tmp_root / "script_build.py"
build_src.write_text(script_text, encoding="utf-8")
print(f"  Archivo a empaquetar: {build_src.name}")

# ==================== PASO 4: Construir .exe único ====================
print("\n[4/4] Construyendo .exe con PyInstaller del venv limpio...")
print("  (debería tardar 2-4 minutos)\n")

hidden_imports = [
    "bs4",
    "lxml",
    "lxml.etree",
    "openpyxl",
    "openpyxl.cell._writer",
    "openpyxl.styles",
    "openpyxl.styles.fills",
    "openpyxl.styles.fonts",
    "openpyxl.utils",
    "openpyxl.worksheet.table",
    "requests",
    "requests.adapters",
    "requests.auth",
    "requests.cookies",
    "requests.models",
    "requests.sessions",
    "urllib3",
    "urllib3.util.retry",
    "urllib3.util.ssl_",
    "certifi",
    "charset_normalizer",
    "idna",
    "selenium.webdriver.chrome",
    "selenium.webdriver.chrome.webdriver",
    "selenium.webdriver.chrome.service",
    "selenium.webdriver.chrome.options",
    "selenium.webdriver.remote.webdriver",
    "selenium.webdriver.support.ui",
    "selenium.webdriver.support.expected_conditions",
    "selenium.common.exceptions",
    "numpy",
    "pandas",
    "pandas.core.arrays",
    "tkinter",
    "tkinter.ttk",
    "tkinter.messagebox",
    "tkinter.filedialog",
    "tkinter.scrolledtext",
]

pyinstaller_exe = (
    venv_dir / "Scripts" / "pyinstaller.exe"
    if os.name == "nt"
    else venv_dir / "bin" / "pyinstaller"
)

cmd = [
    str(pyinstaller_exe),
    "--noconfirm",
    "--onefile",
    "--noconsole",
    "--name",      APP_NAME,
    "--distpath",  str(tmp_dist),
    "--workpath",  str(tmp_build),
    "--specpath",  str(tmp_root),
]

for hi in hidden_imports:
    cmd += ["--hidden-import", hi]

cmd.append(str(build_src))

t0 = time.time()
result = subprocess.run(cmd, capture_output=True, text=True, cwd=str(tmp_root))
elapsed = time.time() - t0
print(f"  PyInstaller tardó {elapsed/60:.1f} minutos.")

if result.returncode != 0:
    print("  ERROR en PyInstaller. Mensajes relevantes:")
    lines  = (result.stdout + result.stderr).splitlines()
    errors = [l for l in lines if any(
        k in l.lower() for k in ["error", "failed", "modulenotfound", "importerror", "permissionerror"]
    )]
    for l in errors[-30:]:
        print("   ", l)
    shutil.rmtree(tmp_root, ignore_errors=True)
    raise RuntimeError("PyInstaller falló. Revisa los mensajes anteriores.")

print("  PyInstaller completado exitosamente.")

exe_name = f"{APP_NAME}.exe" if os.name == "nt" else APP_NAME
exe_src  = tmp_dist / exe_name
exe_dst  = root / exe_name

if exe_src.exists():
    if exe_dst.exists():
        os.remove(exe_dst)
    shutil.move(str(exe_src), str(exe_dst))
    size_mb = exe_dst.stat().st_size / (1024 * 1024)
    print(f"\n  Ejecutable listo: {exe_dst}")
    print(f"  Tamaño: {size_mb:.1f} MB")
else:
    print(f"  ADVERTENCIA: no se encontró {exe_name} en {tmp_dist}")

shutil.rmtree(tmp_root, ignore_errors=True)
print("  Carpeta temporal eliminada.")

print("\n" + "=" * 60)
print(f"  LISTO. Ejecuta {exe_name} directamente.")
print(f"  Requiere Chrome instalado en el equipo.")
print("=" * 60)