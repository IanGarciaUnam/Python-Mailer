# -*- coding: utf-8 -*-

from distutils.core import setup 
import py2exe 
 
setup(name="Mailer", 
 version="1.0", 
 description="Python-Mailer", 
 author="Ian Garcia", 
 author_email="iangarcia@ciencias.unam.mx", 
 url="https://github.com/IanGarciaUnam/Python-Mailer", 
 license="GPL", 
 scripts=["main.py, 
 console=["main.py, 
 options={"py2exe": {"bundle_files": 1}}, 
 zipfile=None,
 )