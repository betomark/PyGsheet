from setuptools import setup
from codecs import open
from os import path

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='pygsheet',

    version='0.1.19b18',

    description='Wrapper for Google sheets API',
    long_description=long_description,

    url='https://github.com/betomark/PyGsheet',

    author='Alberto Marquez Alarcon',
    author_email='almark16@gmail.com',

    license='GPL-3.0',

    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Intended Audience :: Information Technology',
        'License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)',
        'Natural Language :: English',
        'Operating System :: POSIX :: Linux',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: MacOS',
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.0'
        'Programming Language :: Python :: 3.1',
        'Programming Language :: Python :: 3.2',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Topic :: Office/Business',
        'Topic :: Office/Business :: Office Suites',
        'Topic :: Software Development',
        'Topic :: Scientific/Engineering :: Information Analysis',
    ],

    keywords='spreadsheet Google',
    packages=['pygsheet'],
    install_requires=[
            'oauth2client',
            'webcolors',
            'httplib2',
            'google-api-python-client'
    ],
)
