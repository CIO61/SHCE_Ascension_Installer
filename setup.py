from distutils.core import setup
import py2exe

setup(console=[
    {
        "script": "BootstrapInstaller.py",
        "icon_resources": [(0, "icon.ico")]
    },
    {
        "script": "configure_custom_graphics.py",
        "icon_resources": [(0, "icon.ico")]
    }
])
