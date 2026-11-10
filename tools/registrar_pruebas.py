"""Corre las pruebas y guarda el resultado en data/verificacion_pruebas.json, que muestra la app.

Ejecutar desde la raíz del repositorio:  python tools/registrar_pruebas.py
"""
import json
import subprocess
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
XML = RAIZ / "data" / "_pytest.xml"

r = subprocess.run([sys.executable, "-m", "pytest", "-q", "--no-header", f"--junitxml={XML}", str(RAIZ / "tests")],
                   cwd=RAIZ, capture_output=True, text=True)
pruebas = []
for caso in ET.parse(XML).getroot().iter("testcase"):
    estado = "falló" if caso.find("failure") is not None or caso.find("error") is not None else (
        "omitida" if caso.find("skipped") is not None else "pasó")
    pruebas.append({"archivo": caso.get("classname", "").split(".")[-1] + ".py", "prueba": caso.get("name"),
                    "estado": estado, "segundos": round(float(caso.get("time", 0)), 3)})
XML.unlink()
salida = {"fecha": datetime.now().strftime("%Y-%m-%d %H:%M"), "total": len(pruebas),
          "pasaron": sum(p["estado"] == "pasó" for p in pruebas), "pruebas": pruebas}
(RAIZ / "data" / "verificacion_pruebas.json").write_text(json.dumps(salida, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"{salida['pasaron']}/{salida['total']} pasaron")
sys.exit(r.returncode)
