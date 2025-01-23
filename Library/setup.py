from setuptools import setup, find_packages

setup(
    name='easyconvert',
    version='0.0.11a',
    packages=find_packages(),
    install_requires=[
        'Pillow',
    ],
    author='NikitaFanK12',
    author_email='',
    description='With this library, you can convert files from one type to another',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/NikitaFanK12/EasyConvert',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
)
