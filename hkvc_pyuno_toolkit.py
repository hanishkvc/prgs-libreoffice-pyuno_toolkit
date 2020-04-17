#!/usr/bin/env python3
# A simple pyuno helper toolkit/library
# v20200418IST0157, HanishKVC
#

import os
import uno
import time


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


def oo_opendoc(desktop, filePath):
    document = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(filePath)), "_blank", 0, ())
    return document


def oo_getsheets(document):
    controller = document.getCurrentController()
    sheets = document.getSheets()
    return sheets, controller


if __name__ == "__main__":
    oo_run()
    time.sleep(2)
    oo = oo_connect()
    doc = oo_opendoc(oo, "/tmp/t.xlsx")
    sheets, ctlr = oo_getsheets(doc)
    print("NumRows", len(sheets[0].getRows()), "\n Rows", sheets[0].getRows())
    print("NumCols", len(sheets[0].getColumns()), "\n Cols", sheets[0].getColumns())
    for sheet in sheets:
        print("Sheet", dir(sheet))
        # sheet.NamedRanges, sheet.getRows(), sheet.Rows, sheet.getColumns
        # sheet.NamedRanges['MyRange'], sheet.NamedRanges.getByName('MyRange')
        # sheet.getPrintAreas(), sheet.showDetail(?)
        print("Cell", dir(sheet.getCellByPosition(5,5)))
        print(sheet.getCellByPosition(5,5).getValue())

