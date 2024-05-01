# excel2pycl

[![Hatch project](https://img.shields.io/badge/%F0%9F%A5%9A-Hatch-4051b5.svg)](https://github.com/pypa/hatch)


# Installing environment

Clone the repository.

```bash

    git clone https://github.com/esoft-tech/py-bc-excel2pycl.git
```

Install hatch using the guide from their website – https://hatch.pypa.io/latest/install/.

Inside project directory install dependencies.

```bash
    hatch env create test
```

After that, install pre-commit hook.

```bash
    hatch env run -e test pre-commit install
```

If you can't find the path of venv (or vscode can't find), you can use this command:

```bash
    hatch env find
```

Done 🪄 🐈‍⬛ Now you can develop.

# If you want contributing

- Check that ruff passed.
- Check that mypy passed.
- Before adding or changing the functionality, write unit tests.
- Check that unit tests passed and .