import sys, os
from setuptools import setup, Extension

os.chdir(os.path.dirname(__file__))

USE_CYTHON = True
"""
Specify whether to use Cython to build the extensions or use the C files (that were previously generated with Cython):

- Set it to `True` to enable building extensions using Cython.
- Set it to `False` to build extensions from the C files (that were previously generated with Cython).
- Set it to `auto` to build with Cython if available, otherwise from the C file.
"""

if USE_CYTHON:
    try:
        from Cython.Distutils import build_ext
    except ImportError:
        if USE_CYTHON=="auto":
            USE_CYTHON=False
        else:
            raise

cmdclass = { }
"""Dictionary of commands to pass to setuptools.setup()"""

ext_modules = [ ]
"""List of extension modules to pass to setuptools.setup()"""

if sys.version_info[0] == 2:
    raise Exception("Python 2.x is not supported")

if USE_CYTHON:
    ext_modules += [
        Extension("macropy.cythonExtensions.common",  ["src/_cythonExtensions/common/common.pyx"]),
        Extension("macropy.cythonExtensions.eventListeners",  ["src/_cythonExtensions/eventListeners/eventListeners.pyx"]),
        Extension("macropy.cythonExtensions.explorerHelper",  ["src/_cythonExtensions/explorerHelper/explorerHelper.pyx"]),
        Extension("macropy.cythonExtensions.imageInverter",  ["src/_cythonExtensions/imageInverter/imageInverter.pyx"]),
        Extension("macropy.cythonExtensions.keyboardHelper",  ["src/_cythonExtensions/keyboardHelper/keyboardHelper.pyx"]),
        Extension("macropy.cythonExtensions.scriptController",  ["src/_cythonExtensions/scriptController/scriptController.pyx"]),
        Extension("macropy.cythonExtensions.systemHelper",  ["src/_cythonExtensions/systemHelper/systemHelper.pyx"]),
        Extension("macropy.cythonExtensions.windowHelper",  ["src/_cythonExtensions/windowHelper/windowHelper.pyx"])
    ]
    cmdclass.update({ "build_ext": build_ext })
else:
    ext_modules += [
        Extension("macropy.cythonExtensions.common",  ["src/_cythonExtensions/common/common.c"]),
        Extension("macropy.cythonExtensions.eventListeners",  ["src/_cythonExtensions/eventListeners/eventListeners.c"]),
        Extension("macropy.cythonExtensions.explorerHelper",  ["src/_cythonExtensions/explorerHelper/explorerHelper.c"]),
        Extension("macropy.cythonExtensions.imageInverter",  ["src/_cythonExtensions/imageInverter/imageInverter.c"]),
        Extension("macropy.cythonExtensions.keyboardHelper",  ["src/_cythonExtensions/keyboardHelper/keyboardHelper.c"]),
        Extension("macropy.cythonExtensions.scriptController",  ["src/_cythonExtensions/scriptController/scriptController.c" ]),
        Extension("macropy.cythonExtensions.systemHelper",  ["src/_cythonExtensions/systemHelper/systemHelper.c"]),
        Extension("macropy.cythonExtensions.windowHelper",  ["src/_cythonExtensions/windowHelper/windowHelper.c"])
    ]

requirements = [ ]
"""List of requirements to pass to setuptools.setup()"""

with open("requirements.txt", "r") as f:
    for line in f.read().split():
        requirements.append(line.split(">=")[0])

# https://stackoverflow.com/questions/58533084/what-keyword-arguments-does-setuptools-setup-accept
setup(
    name="kb_macropy",
    version="1.1.1",
    description="Keyboard listener and automation script.",
    author="Ahmed Tarek",
    author_email="ahmedtarek4377@gmail.com",
    url="https://github.com/Ryen-042/macropy",
    package_dir={"macropy": "src"},
    packages=["macropy"],
    package_data={"macropy": ["Images/static/*", "SFX/*"],},
    cmdclass = cmdclass,
    ext_modules=ext_modules,
    long_description=open("README.md").read(),
    entry_points ={"console_scripts": ["macropy = macropy.__main__:main"]},
    long_description_content_type="text/markdown",
    license="MIT",
    install_requires=requirements,
    zip_safe = False,
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Environment :: Console",
        "Environment :: Win32 (MS Windows)",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Cython",
        "Topic :: Utilities",
    ],
    keywords="keyboard automation script",
)
