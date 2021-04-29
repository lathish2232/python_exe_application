import sys

from cx_Freeze import setup, Executable

setup( name = "Unificater", version = "1.0",

       description = "Unificater web scraping",

       executables = [Executable(script="Unificater.py",icon="logo.ico",
                                 base = "Win32GUI")])
