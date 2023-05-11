from setuptools import setup

setup(
    name='excel2pycl',
    version='1.0.4',
    description='Package for translating an excel file to python code',
    url='https://github.com/etagi-esoft/py-excel-to-func',
    author='Esoft',
    author_email='it@esoft.tech',
    license='MIT',
    packages=['excel2pycl', 'excel2pycl.src', 'excel2pycl.src.utilities', 'excel2pycl.src.translators',
              'excel2pycl.src.tokens', 'excel2pycl.src.tokens.composite_tokens', 'excel2pycl.src.tokens.regexp_tokens',
              'excel2pycl.src.exceptions'],
    install_requires=['openpyxl>=3.1.2'],
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Environment :: Console',
        'Environment :: MacOS X',
        'Environment :: Web Environment',
        'Environment :: Win32 (MS Windows)',
        'License :: OSI Approved :: MIT License',
        'Operating System :: MacOS',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: POSIX :: Linux',
    ],
)
