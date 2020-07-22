from distutils.core import setup
import py2exe

setup(
    console = [{'script': 'arc_download_assign_error.py'}],
    options = {
        'py2exe': {
            'includes': 'decimal',
            },
        }
)