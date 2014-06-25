from cx_Freeze import setup, Executable



setup(
    name = "OneNote Email Notifications",
    version = "0.1",
    description = "An email notifier for OneNote",

    executables = [Executable("console.py")]
    )