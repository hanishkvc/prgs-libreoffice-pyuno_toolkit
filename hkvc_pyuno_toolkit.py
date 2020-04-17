#!/usr/bin/env python3
# A simple pyuno helper toolkit/library
# v20200418IST0227, HanishKVC
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


def oo_test():
    oo_run()
    time.sleep(2)
    oo = oo_connect()
    doc = oo_opendoc(oo, "/tmp/t.xlsx")
    sheets, ctlr = oo_getsheets(doc)
    # sheet.NamedRanges, sheet.getRows(), sheet.Rows, sheet.getColumns
    # sheet.NamedRanges['MyRange'], sheet.NamedRanges.getByName('MyRange')
    # sheet.getPrintAreas(), sheet.showDetail(?)
    #print(dir(sheet))
    #print(dir(sheet.getCellByPosition(c,r)))
    CellContentTypeEMPTY = uno.Enum("com.sun.star.table.CellContentType","EMPTY")
    for sheet in sheets:
        numRows = len(sheet.Rows)
        numCols = len(sheet.Columns)
        for r in range(numRows):
            if r > numRows:
                continue
            print("INFO:NR:{}, NC:{}, R:{}".format(numRows, numCols, r))
            iEmptyCols = 0
            for c in range(numCols):
                if c > numCols:
                    continue
                if (sheet.getCellByPosition(c,r).getType() == CellContentTypeEMPTY):
                    iEmptyCols += 1
                    if iEmptyCols > 10:
                        numCols = c
                        print("INFO:AdjustNumCols:{}:too many EmptyCols, curCol {}".format(numCols, c))
                print("{}\t".format(sheet.getCellByPosition(c,r).getString()), end="")
            print("")
    return doc, sheets, ctlr



if __name__ == "__main__":
    oo_test()

