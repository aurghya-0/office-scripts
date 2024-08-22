from setuptools import setup, find_packages

setup(
    name='office_scripts',
    version='1.0.0',
    description='A collection of office automation scripts for Excel and Google Drive',
    author='Aurghyadip Kundu',
    author_email='adkundu@gmail.com',
    packages=find_packages(),
    install_requires=[
        'pandas',
        'numpy',
        'openpyxl',
        'pydrive'
    ],
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)
