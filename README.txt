CMB Automiser Ver 2.0
=====================


This is a CPWD billing application using computerised measurement bo
oks. CMB Automiser has been developed in hope of being useful and as
time saver for CPWD staff and contractors working in CPWD. Workflow 
is divided into three interfaces - Schedule View, Measurement View a
nd Bill View. Schedule view implements an interface to input the agr
eement schedule/import the schedule from an ods file. Measurement Vi
ew allows to input/manipulate details of CMBs and includes a number 
of convinience measurement item patterns. The Bill View module allow
s billing of these measurement items and implements part rates, exce
ss rates for deviated items, custom previous bill support etc.

Both CMBs and Bills can be rendered to pdf from the respective modul
es. In addition Bill View also exports final bill variables into an 
ods file for further processing.


Installation
------------

**CMB Automiser** can be downloaded from `https://github.com/manuvar
key/cmbautomiser/` and installed using ``python setup.py install``. 
It has been tested with Python 2.7 and TeX Live 2014, and has the 
following extra dependencies.

Python Modules:
undo - included along with application.
ezodf - included along with application.
libxml - Not included
dill - Not included

Latex Modules:
xr-hyper, hyperref, longtable, tabu, lastpage, geometry, hyphenat, 
xstring, forloop, fancyhdr
