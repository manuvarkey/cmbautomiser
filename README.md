
# CMB Automiser

CMB Automiser is a computerised quantity accounting and cost billing application for use in public works organisations. It allows the user to perform these objectives with an intuitive interface and logical work flow. The results are presented in fully formatted, cross referenced pdf documents.

The program is organised in three tabs/views - schedule, measurement and bill views. Schedule view implements an interface to input the agreement schedule/import the schedule from a .xlsx file. Measurement view allows input/manipulation of details of CMBs and includes a number of convenient measurement item patterns. The Bill view module allows billing of selected measurement items and implements part rates, excess rates for deviated items, custom previous bill support etc.

Both CMBs and Bills can be rendered into a pdf /.xlsx document from the respective modules. In addition Bill view also allows the generation of the deviation statement as a .xlsx file.

## Installation

**CMB Automiser** binaries can be downloaded from the following website
[CMB Automiser](http://manuiisc.blogspot.in/p/blog-page.html)

Source code can be downloaded from the following website
[GitHub](https://github.com/manuvarkey/cmbautomiser/)

### Source installation

Application can be installed using `python setup.py install`. It has been tested with Python 3.4 and TeX Live 2014, and has the following extra dependencies.

#### Python Modules

*undo* - Included along with distribution.
*openpyxl* - Included along with distribution.
*jdcal* - Included along with distribution.
*libxml* - Not included

#### Latex Modules

*xr-hyper, hyperref, longtable, tabu, lastpage, geometry, hyphenat, xstring, forloop, fancyhdr*
