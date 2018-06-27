from distutils.core import setup
import py2exe

setup(
    console = [{'script': 'arc_remove_error.py'}],
    options = {
        'py2exe': {
            'includes': 'decimal',
            },
        }
)
