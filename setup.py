import sys, os
import glob
from setuptools import setup, Extension
from Cython.Build import cythonize

os.chdir(os.path.dirname(__file__))

# Forcing all the source files to be recompiled.
os.environ["CYTHON_FORCE_REGEN"] = "1" if "--force" in sys.argv else "0"


ENABLE_PROFILING = False
"""
Specify whether to enable profiling or not. Note that profiling cause a slight overhead to each function call.
See https://cython.readthedocs.io/en/latest/src/tutorial/profiling_tutorial.html for more information.
"""

if "--profile" in sys.argv:
    ENABLE_PROFILING = True
    sys.argv.remove("--profile")

USE_CYTHON = 1
"""
Specify whether to use `Cython` to build the extensions from the `.pxd` files or use the already generated `.c` files:

- Set it to `1` to enable building extensions using Cython.
- Set it to `0` to build extensions from the `.c` files.
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

packages = ["macropy", "macropy.cythonExtensions"]

if sys.version_info[0] == 2:
    raise Exception("Python 2.x is not supported")

if USE_CYTHON:
    cython_extensions = glob.glob("src/cythonExtensions/**/*.pyx", recursive=True)
    cmdclass.update({ "build_ext": build_ext })
else:
    cython_extensions = glob.glob("src/cythonExtensions/**/*.c", recursive=True)

for extension_path in cython_extensions:
    ext_filename = os.path.splitext(os.path.basename(extension_path))
    
    packages.append("macropy.cythonExtensions." + ext_filename[0])
    
    ext_modules.append(Extension(
        name = f"macropy.cythonExtensions.{ext_filename[0]}.{ext_filename[0]}",
        sources = [f"src/cythonExtensions/{ext_filename[0]}/{ext_filename[0]}{ext_filename[1]}"],
        define_macros = [("CYTHON_TRACE", "1")] if ENABLE_PROFILING else [ ],
    ))

# https://stackoverflow.com/questions/58533084/what-keyword-arguments-does-setuptools-setup-accept
setup(
    packages=packages,
    cmdclass = cmdclass,
    ext_modules=cythonize(ext_modules, language_level=3)
)
