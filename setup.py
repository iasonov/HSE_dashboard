from setuptools import setup, find_packages

setup(
    name="HSE_dashboard",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "requests",
        "gspread",
        "oauth2client",
    ],
)
