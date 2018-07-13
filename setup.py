from distutils.core import setup
import py2exe

setup(
    console = [{'script': 'arc_remove_error_retail.py'}],
    options = {
        'py2exe': {
            'includes': 'decimal',
            },
        }
)