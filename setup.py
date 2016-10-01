from setuptools import setup, find_packages 

setup(
    # Application name:
    name="CMBAutomiser3",

    # Version number (initial):
    version="3.1.2",

    # Application author details:
    author="Manu Varkey",
    author_email="manuvarkey@gmail.com",

    # Packages
    packages = find_packages(),
    include_package_data = True, # Include additional files into the package
    data_files=[('/usr/share/applications', ['applications/cmbautomiser3.desktop']),
                  ('/usr/share/pixmaps', ['pixmaps/cmbautomiser3.png']),
                  ('/usr/bin', ['bin/cmbautomiser3'])],


    # Details
    maintainer="Manu Varkey",
    maintainer_email="manuvarkey@gmail.com",
    url="http://manuvarkey.github.io/cmbautomiser/",
    license="GPL-3.0",
    description="Computerised quantity accounting and cost billing application for use in public works organisations.",

    long_description=  "CMB Automiser is a computerised quantity accounting and cost billing application for use in public works organisations. It allows the user to perform these objectives with an intuitive interface and logical work flow. The results are presented in fully formatted, cross referenced pdf documents.",
    
     classifiers=[
          'Development Status :: 5 - Production/Stable',
          'Intended Audience :: End Users/Desktop',
          'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
          'Operating System :: Microsoft :: Windows',
          'Operating System :: POSIX',
          'Programming Language :: Python',
          'Topic :: Office/Business',
          ],
    )
