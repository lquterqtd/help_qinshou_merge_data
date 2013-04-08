__author__ = 'Administrator'
#coding:cp936
from distutils.core import setup
import py2exe
options = {
    "py2exe":
        {
            "compressed": 1,
            "optimize": 2,
            "bundle_files": 1
        }
}
setup(
    options=options,
    zipfile=None,
    console=[{
        "script": "main.py",
        "icon_resources": [(
            1, "logo.ico"
        )]
    }]
)