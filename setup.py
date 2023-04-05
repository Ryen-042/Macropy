USE_CYTHON = True
"""
Specify whether to use Cython to build the extensions or use the C files (that were previously generated with Cython):

- Set this to True to enable building extensions using Cython.
- Set it to False to build extensions from the C files (that were previously generated with Cython).
- Set it to 'auto' to build with Cython if available, otherwise from the C file.
"""
import sys
from setuptools import setup, Extension

if USE_CYTHON:
    try:
        from Cython.Distutils import build_ext
    except ImportError:
        if USE_CYTHON=='auto':
            USE_CYTHON=False
        else:
            raise

cmdclass = { }
"""Dictionary of commands to pass to setuptools.setup()"""

ext_modules = [ ]
"""List of extension modules to pass to setuptools.setup()"""

if sys.version_info[0] == 2:
    raise Exception('Python 2.x is not supported')


if USE_CYTHON:
    ext_modules += [
        Extension("macropy.common",  ["macropy/cython_extensions/common/common.pyx"]),
        Extension("macropy.eventListeners",  ["macropy/cython_extensions/eventListeners/eventListeners.pyx"]),
        Extension("macropy.explorerHelper",  ["macropy/cython_extensions/explorerHelper/explorerHelper.pyx"]),
        Extension("macropy.imageInverter",  ["macropy/cython_extensions/imageInverter/imageInverter.pyx"]),
        Extension("macropy.keyboardHelper",  ["macropy/cython_extensions/keyboardHelper/keyboardHelper.pyx"]),
        Extension("macropy.scriptController",  ["macropy/cython_extensions/scriptController/scriptController.pyx"]),
        Extension("macropy.systemHelper",  ["macropy/cython_extensions/systemHelper/systemHelper.pyx"]),
        Extension("macropy.windowHelper",  ["macropy/cython_extensions/windowHelper/windowHelper.pyx"]),
        
    ]
    cmdclass.update({ 'build_ext': build_ext })
else:
    ext_modules += [
        Extension("macropy.common",  ["macropy/cython_extensions/common/common.c"]),
        Extension("macropy.eventListeners",  ["macropy/cython_extensions/eventListeners/eventListeners.c"]),
        Extension("macropy.explorerHelper",  ["macropy/cython_extensions/explorerHelper/explorerHelper.c"]),
        Extension("macropy.imageInverter",  ["macropy/cython_extensions/imageInverter/imageInverter.c"]),
        Extension("macropy.keyboardHelper",  ["macropy/cython_extensions/keyboardHelper/keyboardHelper.c"]),
        Extension("macropy.scriptController",  ["macropy/cython_extensions/scriptController/scriptController.c" ]),
        Extension("macropy.systemHelper",  ["macropy/cython_extensions/systemHelper/systemHelper.c"]),
        Extension("macropy.windowHelper",  ["macropy/cython_extensions/windowHelper/windowHelper.c"])
    ]

requirements = [ ]
"""List of requirements to pass to setuptools.setup()"""

with open('requirements.txt', 'r') as f:
    for line in f.read().split():
        requirements.append(line.split(">=")[0])

setup(
    name="macropy",
    version="0.0.17",
    description="Keyboard listener and automation script.",
    author="Ahmed Tarek",
    author_email="ahmedtarek4377@gmail.com",
    url="https://github.com/Ryen-042/macropy",
    packages=["macropy", "macropy.common", "macropy.eventListeners", "macropy.explorerHelper", "macropy.imageInverter", 
              "macropy.keyboardHelper", "macropy.scriptController", "macropy.systemHelper", "macropy.windowHelper"],
    package_data={"stub":  ["macropy/common.pyi",
                            "macropy/eventListeners.pyi",
                            "macropy/explorerHelper.pyi",
                            "macropy/imageInverter.pyi",
                            "macropy/keyboardHelper.pyi",
                            "macropy/scriptController.pyi",
                            "macropy/systemHelper.pyi",
                            "macropy/windowHelper.pyi"]},
    cmdclass = cmdclass,
    ext_modules=ext_modules,
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    license="MIT",
    install_requires=requirements,
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        "Environment :: Console",
        "Environment :: Win32 (MS Windows)",
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        "Programming Language :: Python :: 3.10",
        'Programming Language :: Cython',
        "Topic :: Utilities",
    ],
    keywords='keyboard automation script',
)
