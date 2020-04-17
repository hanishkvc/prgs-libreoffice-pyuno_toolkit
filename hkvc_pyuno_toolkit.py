#!/usr/bin/env python3
# A simple pyuno helper toolkit/library
# v20200418IST0100, HanishKVC
#

import os
import uno


def oo_run():
    if os.fork() == 0:
        os.system('libreoffice --accept="socket,host=localhost,port=2002;urp;&"')
        exit()
    else:
        print("INFO:oo_run:libreoffice should be started now")


def oo_connect():
    localCtxt = uno.getComponentContext()
    resolver = localCtxt.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localCtxt)
    ctxt = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    smgr = ctxt.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctxt)
    PropertyValue = smgr.createInstanceWithContext("com.sun.star.beans.PropertyValue", ctxt)
    return desktop


oo_run()
oo = oo_connect()
print(dir(oo))

