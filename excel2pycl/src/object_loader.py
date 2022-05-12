import importlib.util


def load_module(module_path: str):
    spec = importlib.util.spec_from_file_location(module_path, module_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    return module
