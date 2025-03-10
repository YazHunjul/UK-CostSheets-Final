from setuptools import setup, find_packages

setup(
    name="uk-costsheets",
    packages=find_packages(),
    include_package_data=True,
    package_data={
        'app.costSheetGen.costSheetResources': ['*.xlsx', '*.docx'],
    },
) 