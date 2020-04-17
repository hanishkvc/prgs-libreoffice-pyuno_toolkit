#!/usr/bin/env python3
# A simple pyuno helper toolkit/library
# v20200418IST0255, HanishKVC
#

import os
import uno
import time
import sys



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


def _oo_props(**args):
    props = []
    for key in args:
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = key
        prop.Value = args[key]
        props.append(prop)
    return tuple(props)


def oo_savedoc(doc, filePath, filterName):
    props = _oo_props(FilterName=filterName)
    doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(filePath)), props)


def oo_test(oo):
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
            if (r%100) == 0:
                print("INFO:NR:{}, NC:{}, R:{}".format(numRows, numCols, r), file=sys.stderr)
            iEmptyCols = 0
            for c in range(numCols):
                if c > numCols:
                    continue
                if (sheet.getCellByPosition(c,r).getType() == CellContentTypeEMPTY):
                    iEmptyCols += 1
                    if iEmptyCols > 10:
                        numCols = c
                        print("INFO:AdjustNumCols:{}:too many EmptyCols, curCol {}".format(numCols, c))
                else:
                    # Ensure that we look for atleast 10 cols beyond current row's first empty col
                    if numCols < (c+10):
                        numCols = c+10
                        print("INFO:AdjustNumCols:{}:too few EmptyCols buffer, curCol {}".format(numCols, c))
                print("{}\t".format(sheet.getCellByPosition(c,r).getString()), end="")
            print("")
    return doc, sheets, ctlr


def oo_conv_ss2csv(oo, sIn, sOut):
    doc = oo_opendoc(oo, sIn)
    oo_savedoc(doc, sOut, filterName="Text - txt - csv (StarCalc)")


if __name__ == "__main__":
    oo_run()
    time.sleep(2)
    oo = oo_connect()
    # python3 hkvc_pyuno_convert.py ss2csv /tmp/t.xlsx /tmp/t.csv
    if sys.argv[1] == "ss2csv":
        oo_conv_ss2csv(oo, sys.argv[2], sys.argv[3])
    else:
        oo_test(oo)

