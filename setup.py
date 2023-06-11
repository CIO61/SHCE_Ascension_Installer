import py2exe

py2exe.freeze(
    console=[{"script": "BootstrapInstaller.py"},
             {"script": "configure_custom_graphics.py"},
             {"script": "selfupdate.py"}]
)
