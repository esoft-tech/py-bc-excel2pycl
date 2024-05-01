import importlib.util
from types import ModuleType


def load_module(module_path: str) -> ModuleType:
    spec = importlib.util.spec_from_file_location(module_path, module_path)
    if spec is None:
        raise Exception("Spec has not found", module_path)

    module = importlib.util.module_from_spec(spec)

    if module is None:
        raise Exception("Module is None", module_path)

    loader = spec.loader

    if loader is None:
        raise Exception("Loader is None", module_path)

    loader.exec_module(module)

    return module
