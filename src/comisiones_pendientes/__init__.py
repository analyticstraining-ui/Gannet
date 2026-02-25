import importlib.util
import os

_path = os.path.join(os.path.dirname(__file__), "Comisiones Pendientes Prov.py")
_spec = importlib.util.spec_from_file_location("comisiones_pendientes_prov", _path)
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)

generate_comisiones_report = _mod.generate_comisiones_report

__all__ = ["generate_comisiones_report"]
