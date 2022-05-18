from setuptools import setup

setup(
    name='excel2pycl',
    version='0.0.9',
    description='Package for translating an excel file to python code',
    url='https://github.com/etagi-esoft/py-excel-to-func',
    author='Esoft',
    author_email='it@esoft.tech',
    license='MIT',
    packages=['excel2pycl', 'excel2pycl.src'],
    install_requires=['openpyxl>=3.0.9'],

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
