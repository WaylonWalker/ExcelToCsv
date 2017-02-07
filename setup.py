from setuptools import setup, find_packages


with open('requirements.txt') as f:
    required = f.read().splitlines()

with open('README.md') as f:
    long_description = f.read()


setup(name='ExcelToCsv',
      version='0.1',
      description=long_description,
      url='https://github.com/WaylonWalker/ExcelToCsv',
      author='WaylonWalker',
      author_email='quadmx08@gmail.com',
      license='MIT',
      packages=find_packages(),
      install_requires=required,
      zip_safe=False,
      entry_points={
        'console_scripts': ['ExcelToCsv=ExcelToCsv:console_tool'],
                    }

        )
