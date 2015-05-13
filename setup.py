from distutils.core import setup

setup(
    # Application name:
    name="CMBAutomiser",

    # Version number (initial):
    version="2.0.0",

    # Application author details:
    author="Manu Varkey",
    author_email="manuvarkey@gmail.com",

    # Packages
    packages=["cmbautomiser"],
    package_dir={'cmbautomiser': 'cmbautomiser'},
    package_data={'cmbautomiser': ['latex/*','ezodf/*','interface/*','templates/*']},

    # Include additional files into the package
    include_package_data=True,
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
                        
    # Dependent packages (distributions)
    # install_requires=['texlive-collection-basic', 'texlive-collection-latex', 'texlive-collection-latexextra',
    #       'texlive-collection-latexrecommended' ,'texlive-collection-fontsrecommended', 
    #       'texlive-collection-generic', 'texlive-collection-genericextra', 'texlive-collection-langenglish', 
    #       'texlive-hyperref', 'texlive-hyper', 'texlive-hyphenat', 'texlive-tabu', 'texlive-lastpage',
    #       'texlive-geometry', 'texlive-xstring', 'texlive-forloop', 'texlive-fancyhdr'],
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
