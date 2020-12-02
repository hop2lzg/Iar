from distutils.core import setup
import py2exe

setup(
    console = [{'script': 'arc_update_commission_user_sync.py'}],
    options = {
        'py2exe': {
            'includes': 'decimal',
            },
        }
)