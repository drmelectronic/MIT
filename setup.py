#*********************************************
#        Auto-Generated With py2Nsis
#*********************************************

import warnings 
#ignore the sets DeprecationWarning
warnings.simplefilter('ignore', DeprecationWarning) 
import py2exe
warnings.resetwarnings() 
from distutils.core import setup
		
target = {
'script' : "C:\\Documents and Settings\\Daniel\\Mis documentos\\Proyectos\\MIT03-6\\__init__.py",
'version' : "1.1",
'company_name' : "ECONAIN S.A.C.",
'copyright' : "GPL v3.0",
'name' : "MIT03", 
'dest_base' : "MIT03", 
'icon_resources': [(1, "C:\\Documents and Settings\\Daniel\\Mis documentos\\Proyectos\\MIT03-6\\logo.ico")]
}		



setup(

	data_files = [],
    
    zipfile = None,

	options = {"py2exe": {"compressed": 0, 
						  "optimize": 0,
						  "includes": ['atk', 'pango', 'cairo', 'pangocairo', 'gobject', 'gio'],
						  "excludes": ['_gtkagg', '_tkagg', 'bsddb', 'curses', 'email', 'pywin.debugger', 'pywin.debugger.dbgcon', 'pywin.dialogs', 'tcl', 'Tkconstants', 'Tkinter'],
						  "packages": [],
						  "bundle_files": 3,
						  "dist_dir": "C:\\Documents and Settings\\Daniel\\Mis documentos\\Proyectos\\MIT03-6\\dist",
						  "xref": False,
						  "skip_archive": False,
						  "ascii": False,
						  "custom_boot_script": '',                          
						 }
			  },
	console = [],
	windows = [target],
	service = [],
	com_server = [],
	ctypes_com_server = []
)
		
