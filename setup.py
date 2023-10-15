import py2exe

py2exe.freeze(
    console=[{"script": "AscensionInstaller.py"},
             {"script": "configure_custom_graphics.py"},
             {"script": "selfupdate.py"}]
)
