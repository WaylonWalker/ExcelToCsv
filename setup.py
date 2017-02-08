from setuptools import setup, find_packages


with open('requirements.txt') as f:
    required = f.read().splitlines()

with open('README.rst') as f:
    long_description = f.read()


setup(name='ExcelToCsv',
      version='0.1.1',
      description='A python package to convert excel spreadsheets to csv files',
      long_description=long_description,
      url='https://github.com/WaylonWalker/ExcelToCsv',
      author='Waylon Walker',
      author_email='quadmx08@gmail.com',
      license='MIT',
      packages=find_packages(),
      install_requires=required,
      classifiers=[
        "Development Status :: 3 - Alpha",
        "Topic :: Utilities"],
      zip_safe=False,
      entry_points={
        'console_scripts': ['ExcelToCsv=ExcelToCsv:console_tool'],
                    }
        )
