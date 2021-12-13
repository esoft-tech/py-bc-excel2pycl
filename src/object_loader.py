import re


def load_object(module: str, object_name: str):
    module_name = module
    if module_name.find('.') != -1:
        module_name: str = re.findall(r'^.*\.([\d\w]+)$', module)[0]
    return __import__(module).__dict__[module_name].__dict__[object_name]
