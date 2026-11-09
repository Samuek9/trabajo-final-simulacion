import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from investment_sim.config.parameters import cargar, leer_tabla  # noqa: E402


@pytest.fixture(scope="session")
def tabla():
    return leer_tabla()


@pytest.fixture
def base(tabla):
    return cargar("base", tabla=tabla)


# Un mundo sin aleatoriedad ni costos: todo rinde cero y nada se paga.
QUIETO = {
    "inf.c": -30.0, "inf.phi": 0.0, "inf.sigma": 0.0, "inf.inicial": 1e-9,
    "tpm.inicial": 0.0, "tpm.hace_3m": 0.0, "tpm.min": 0.0, "junta.modelo": "tendencia",
    "junta.bajando": "0;1;0", "junta.estable": "0;1;0", "junta.subiendo": "0;1;0",
    "cta.alfa": 0.0, "imp.retencion": 0.0, "imp.retencion_cuenta": 0.0, "imp.gmf": 0.0,
    "cdt.modelo": "fase", "cdt.fase_subida": "0;0", "cdt.fase_estable": "0;0", "cdt.fase_bajada": "0;0",
    "cdt.margen_min": 0.0, "cdt.margen_max": 0.0,
    "etf.mu": 0.0, "etf.vol_E": 0.0, "etf.vol_N": 0.0, "etf.vol_C": 0.0, "etf.comision": 0.0,
    "val.ar1": "0;0;0", "val.mu_E": 0.0, "val.mu_N": 0.0, "val.mu_C": 0.0, "val.sd_E": 0.0, "val.sd_N": 0.0, "val.sd_C": 0.0,
    "inm.costo_compra": 0.0, "inm.canon_inicial": 0.0, "inm.ph": 0.0, "inm.predial": 0.0,
    "inm.mantenimiento": 0.0, "inm.pintura": 0.0, "inm.costo_venta": 0.0,
    "vac.p_E": 0.0, "vac.p_N": 0.0, "vac.p_C": 0.0,
    "liq.tasa_E": 0.0, "liq.tasa_N": 0.0, "liq.tasa_C": 0.0,
}


@pytest.fixture
def quieto(base):
    return base.con_dict(QUIETO)
