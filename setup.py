from setuptools import setup

setup(name='html2excel',
      version='0.1.0',
      description='Export Html table to Excel file (.xlsx)',
      url='https://github.com/nguyenminhquan/html2excel',
      author='Nguyen Minh Quan',
      author_email='nguyenminhquan2195@gmail.com',
      license='MIT',
      packages=['html2excel'],
      install_requires=[
          'bs4',
          'openpyxl'
      ],
      zip_safe=False)
