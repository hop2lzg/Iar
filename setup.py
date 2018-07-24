from distutils.core import setup
import py2exe

setup(
    console = [{'script': 'arc_put_error.py'}],
    options = {
        'py2exe': {
            'includes': 'decimal',
            },
        }
)