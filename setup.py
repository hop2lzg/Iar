from distutils.core import setup
import py2exe

setup(
    console = [{'script': 'arc_update_commission.py'}],
    options = {
        'py2exe': {
            'includes': 'decimal',
            },
        }
)