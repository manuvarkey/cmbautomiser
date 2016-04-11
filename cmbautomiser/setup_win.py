from distutils.core import setup 
import py2exe 
import sys, os, site, shutil

site_dir = site.getsitepackages()[1] 
include_dll_path = os.path.join(site_dir, "gnome") 

gtk_dirs_to_include = ['etc', 'lib\\gtk-3.0', 'lib\\girepository-1.0', 'lib\\gio', 'lib\\gdk-pixbuf-2.0', 'share\\glib-2.0', 'share\\fonts', 'share\\icons', 'share\\themes\\Default', 'share\\themes\\HighContrast'] 

gtk_dlls = [] 
tmp_dlls = [] 
cdir = os.getcwd() 
for dll in os.listdir(include_dll_path): 
    if dll.lower().endswith('.dll'): 
        gtk_dlls.append(os.path.join(include_dll_path, dll)) 
        tmp_dlls.append(os.path.join(cdir, dll)) 

for dll in gtk_dlls: 
    shutil.copy(dll, cdir) 

# -- change main.py if needed -- #
setup(
    # Application name:
    name="CMBAutomiser",

    # Version number (initial):
    version="2.2.0",

    # Application author details:
    author="Manu Varkey",
    author_email="manuvarkey@gmail.com",
    windows=['__init__.py'],
    data_files = [("Microsoft.VC90.CRT", ['winlibs\\msvcr90.dll'])],
    options={ 
    'py2exe': { 
        'includes' : ['gi','lxml.etree', 'lxml._elementpath', 'gzip'], 
        'packages': ['gi'],
        "skip_archive": True,
              }
            },
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

dest_dir = os.path.join(cdir, 'dist') 

if not os.path.exists(dest_dir):
    os.makedirs(dest_dir)

for dll in tmp_dlls: 
    shutil.copy(dll, dest_dir) 
    os.remove(dll) 

for d  in gtk_dirs_to_include: 
    shutil.copytree(os.path.join(site_dir, "gnome", d), os.path.join(dest_dir, d))
