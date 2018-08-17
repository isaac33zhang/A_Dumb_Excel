from setuptools import setup, find_packages


packages = find_packages('lib')
scripts = ['bin/pangpang.py']
setup(
    name='excel',
    version='1.0',
    packages=packages,
    scripts=scripts,
    package_dir={'': 'lib'},
    author='fanzhaf',
    author_email='isaac33zhang@gmail.com',
    description="Excel interaction package for Pangpang's request"
)
