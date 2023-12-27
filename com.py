from cx_Freeze import Executable, setup

setup(
    name="mail_to_PDF",
    version="1.0",
    description="Description of your script",
    executables=[Executable("app.py")],
)
