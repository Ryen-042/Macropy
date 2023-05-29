import sys, os
import glob
from setuptools import setup, Extension

os.chdir(os.path.dirname(__file__))

# Forcing the cython files to be recompiled regardless of modification times and changes.
os.environ["CYTHON_FORCE_REGEN"] = "1" if "--force" in sys.argv else "0"

ENABLE_PROFILING = "--profile" in sys.argv
"""
Specify whether to enable profiling or not. Note that profiling cause a slight overhead to each function call.
See https://cython.readthedocs.io/en/latest/src/tutorial/profiling_tutorial.html for more information.
"""

USE_CYTHON = 1
"""
Specify whether to use `Cython` to build the extensions or use the `C` files (that were previously generated with Cython):

- Set it to `1` to enable building extensions using Cython.
- Set it to `0` to build extensions from the C files (that were previously generated with Cython).
- Set it to `-1` to build with Cython if available, otherwise from the C file.
"""

# Source: https://stackoverflow.com/questions/28301931/how-to-profile-cython-functions-line-by-line
if ENABLE_PROFILING:
    from Cython.Compiler.Options import get_directive_defaults
    
    directive_defaults = get_directive_defaults()
    directive_defaults['linetrace'] = True
    directive_defaults['binding'] = True

if USE_CYTHON:
    try:
        from Cython.Distutils import build_ext
    except ImportError:
        if USE_CYTHON == -1:
            USE_CYTHON = 0
        else:
            raise

cmdclass = { }
"""Dictionary of commands to pass to setuptools.setup()"""

ext_modules = [ ]
"""List of extension modules to pass to setuptools.setup()"""

if sys.version_info[0] == 2:
    raise Exception("Python 2.x is not supported")

if USE_CYTHON:
    cython_extensions = glob.glob("src/cythonExtensions/**/*.pyx", recursive=True)
    cmdclass.update({ "build_ext": build_ext })
else:
    cython_extensions = glob.glob("src/cythonExtensions/**/*.c", recursive=True)

for extension_path in cython_extensions:
    extension = os.path.splitext(os.path.basename(extension_path))
    if extension[0] == "__init__":
        continue
    
    ext_modules.append(Extension(
        name = f"macropy.cythonExtensions.{extension[0]}.{extension[0]}",
        sources = [f"src/cythonExtensions/{extension[0]}/{extension[0]}{extension[1]}"],
        define_macros = [("CYTHON_TRACE", "1")] if ENABLE_PROFILING else [ ],
    ))

requirements = [ ]
"""List of requirements to pass to setuptools.setup()"""

with open("requirements.txt", "r") as f:
    for line in f.read().split():
        requirements.append(line.split(">=")[0])

# https://stackoverflow.com/questions/58533084/what-keyword-arguments-does-setuptools-setup-accept
setup(
    name="kb_macropy",
    version="1.1.8",
    description="Keyboard listener and automation script.",
    author="Ahmed Tarek",
    author_email="ahmedtarek4377@gmail.com",
    url="https://github.com/Ryen-042/macropy",
    package_dir={"macropy": "src"},
    packages=["macropy"],
    package_data={"macropy": ["Images/static/*", "SFX/*"]},
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
