from distutils.core import setup
import py2exe

setup(
    console = [{'script': 'arc_download_refund.py'}],
    options = {
        'py2exe': {
            'includes': 'decimal',
            },
        }
)