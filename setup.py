from setuptools import setup, find_packages 

setup(
    # Application name:
    name="CMBAutomiser",

    # Version number (initial):
    version="3.0.0",

    # Application author details:
    author="Manu Varkey",
    author_email="manuvarkey@gmail.com",

    # Packages
    packages = find_packages(),
    include_package_data = True, # Include additional files into the package
    data_files=[('/usr/share/applications', ['applications/cmbautomiser.desktop']),
                  ('/usr/share/pixmaps', ['pixmaps/cmbautomiser.png']),
                  ('/usr/bin', ['bin/cmbautomiser'])],


    # Details
    maintainer="Manu Varkey",
    maintainer_email="manuvarkey@gmail.com",
    url="https://github.com/manuvarkey/cmbautomiser",
    license="GPL-3.0",
    description="CPWD billing application using computerised measurement books.",

    long_description=  "CMB Automiser is a CPWD billing application using computerised measurement books. CMB Automiser has been developed in hope of being useful and as time saver for CPWD staff and contractors working in CPWD.",
    
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
