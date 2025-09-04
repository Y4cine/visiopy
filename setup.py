from setuptools import setup, find_packages

setup(
    name='visiopy',
    version='0.2.1',
    packages=find_packages(),
    install_requires=['pywin32'],
    description='A library to automate Visio operations.',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    author='Yacine Gacem',
    author_email='yacine.gacem@gmail.com',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)