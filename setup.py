from setuptools import setup, find_packages

setup(
    name="certificake",
    version='0.0.1',
    packages=find_packages(),
    author="Frederico Schmitt Kremer",
    author_email="fred.s.kremer@gmail.com",
    description="A CLI tool to create event certicates automatically using PPTX templates and lists of participants CSV/XLSX file",
    long_description=open("README.md").read(),
    long_description_content_type='text/markdown',
    keywords="office",
    entry_points = {'console_scripts':[
        'certificake = certificake:main'
        ]},
    install_requires = [
        'pandas',
        'openpyxl',
        'python-pptx',
        'jinja2'
    ]
)