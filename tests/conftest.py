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


from investment_sim.experiments.verification import QUIETO  # noqa: E402


@pytest.fixture
def quieto(base):
    return base.con_dict(QUIETO)
