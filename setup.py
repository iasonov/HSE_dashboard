from setuptools import setup, find_packages

setup(
    name="HSE_dashboard",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        pandas, gspread, oauth2client # Add project dependencies here
    ],
)