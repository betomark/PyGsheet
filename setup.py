from setuptools import setup

setup(name='pygsheet',
      version='0.1',
      description='wrapper for Google sheets API',
      url='https://github.com/betomark/PyGsheet',
      author='Alberto Marquez',
      author_email='almark16@gmail.com',
      license='GPL',
      packages=[],
	  install_requires=[
          'oauth2client',
		  'webcolors',
		  'httplib2',
		  'google-api-python-client'
      ],
      zip_safe=False)