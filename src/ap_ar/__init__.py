import importlib

_mod = importlib.import_module("src.ap_ar.AP&AR")
generate_ap_ar_report = _mod.generate_ap_ar_report

__all__ = ["generate_ap_ar_report"]
