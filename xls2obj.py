from datetime import datetime
import xlrd
from openpyxl import load_workbook
from csv import reader
import json
from pathlib import Path
import os

class XlsObj:
    typd = {
        'str' : lambda v,_ : str(v),
        'float' : lambda v,_ : float(v) if v!='' else None,
        'int' : lambda v,_ : int(float(v)) if v!='' else None,
        'date' : lambda v,s : datetime.strptime(v,s['datefmt']),
        'xldate' : lambda v,_ : v,
        }
    def trim(self,val,triml): return self.trim(val.replace(triml[0],''),triml[1:]) \
        if triml else val
    def __init__(self,row,xos):
        for k,v in xos.fields.items():
            rawval = row[v['col']-1]
            rawvals = rawval.strip() if isinstance(rawval,str) else rawval
            trimval = self.trim(rawvals,v['trim'])
            typfn = self.typd[v['typ']]
            conval = typfn(trimval,v)
            self.__dict__.update({k:conval})

class xls:
    def toval(self,c): return xlrd.xldate.xldate_as_datetime(c.value,self.datemode) \
        if c.ctype == xlrd.XL_CELL_DATE else c.value
    def rows(self,flnm,sheet=0):
        book = xlrd.open_workbook(flnm)
        sh = book.sheets()[sheet]
        self.datemode = book.datemode
        return [ [self.toval(c) for c in r] for r in sh.get_rows() ]

class xlsx:
    def rows(self,flnm,sheet=0): return [
        [ (c.value if c.value!=None else '') for c in r ]
        for r in load_workbook(flnm,read_only=True).get_active_sheet()
        ]

class csv:
    def rows(self,flnm,sheet=0): return reader(open(flnm,errors='ignore'))

class XlsObjs:
    _globals = {
        'typ':'str', 'strtrow':1, 'endpatcol':1, 'endpat':'', 'trim':[], 'remark':'',
        }
    def __iter__(self): return iter(self.objs)
    # Either specfile or specname need to be specified, specfile needs to be a
    # file path. specname is just the basename of the json file and requires
    # XLS2PYSPECDIR env var to be set
    def __init__(self,flnm,specfile=None,specname=None,sheet=0):
        specflnm = specfile if specfile else (
            os.environ['XLS2PYSPECDIR']+ '/' + specname+ '.json'
            )
        spec = json.load(open(specflnm))
        self.g = { **self._globals, **spec.get('globals',{}) }
        self.fields = { k:{ **self.g, **v } for k,v in spec['fields'].items() }
        strtrow = self.g['strtrow'] - 1
        endpatcol = self.g['endpatcol'] - 1
        endpat = self.g['endpat']
        ftyp = globals()[Path(flnm).suffix[1:]]
        self.objs = []
        for i,r in enumerate(ftyp().rows(flnm)):
            if i < strtrow : continue
            if r[endpatcol].strip() == endpat: break
            self.objs = self.objs + [ XlsObj(r,self) ]

