# -*- coding: utf-8 -*-


import os
from distutils.core import setup


# Version numbering follows the conventions of Semantic Versioning 2.0.0:
# 1. MAJOR - changes with backwards-incompatible modifications to the API
# 2. MINOR - for backwards-compatible additions to functionality
# 3. PATCH - for backwards-compatible bug fixes
# Source: http://semver.org/spec/v2.0.0.html
MAJOR = 0
MINOR = 2
PATCH = 0
for_release = False

VERSION = '%d.%d.%d' % (MAJOR, MINOR, PATCH)
if not for_release:
    VERSION += '.dev'

# Write package version.py
with open(os.path.join('officetools', 'version.py'), 'wt') as f:
    version_to_write = '''\
# -*- coding: utf-8 -*-


MAJOR = %d
MINOR = %d
PATCH = %d
VERSION = \'%d.%d.%d'''
    if not for_release:
        version_to_write += '.dev'
    version_to_write += '\'\n'
    f.write(version_to_write % (MAJOR, MINOR, PATCH, MAJOR, MINOR, PATCH))

# Call setup()
setup(
    name='officetools',
    version=VERSION,
    description='Python tools for handling Microsoft Office files',
    long_description='''
officetools
===========
`officetools` is a Python package that provides command-line tools for
handling Microsoft Office files (with file extension `.docx`). In particular,
it provides a script to extract the comments from a Microsoft Word file and
write them to a CSV file.
''',
    license='BSD',
    author='Chris Thoung',
    author_email='chris.thoung@gmail.com',
    url='https://github.com/cthoung/office-tools',
    packages=[
        'officetools',
        'officetools.docx',
        ],
    scripts=[
        'scripts/docx.py',
        ],
    classifiers=[
        'Development Status :: 4 - Beta'
        'Environment :: Console',
        'Intended Audience :: End Users/Desktop',
        'License :: OSI Approved :: BSD License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.2',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Topic :: Office/Business :: Office Suites',
        'Topic :: Utilities',
        ],
    platforms=['Any'],
    )
