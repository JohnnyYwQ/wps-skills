"""Application selection for Task admission; contracts remain application-owned."""

import importlib
from types import MappingProxyType

PROFILES = MappingProxyType({
    "word": MappingProxyType({"create": "createDocument", "open": "openDocument"}),
    "excel": MappingProxyType({"create": "createWorkbook", "open": "openWorkbook"}),
    "ppt": MappingProxyType({"create": "createPresentation", "open": "openPresentation"}),
})


def profile(application):
    if application not in PROFILES:
        raise ValueError("Unsupported Task application: " + str(application))
    return PROFILES[application]


def contracts_for(application):
    profile(application)
    module = importlib.import_module("wps_skills." + application + ".contracts")
    return getattr(module, application.upper() + "_PRODUCTION_CONTRACT_SET")


def compile_request(request, application):
    profile(application)
    # Application entry points allow application-specific admission to evolve.
    module = importlib.import_module("wps_skills." + application + ".task.plan")
    return module.compile_request(request)
