from datetime import datetime
import xlrd
from openpyxl import load_workbook
from csv import reader
import json
from pathlib import Path

class XlsObj:
    typd = {
        'str' : lambda v,_ : str(v),
        'float' : lambda v,_ : float(v) if v!='' else 0,
        'int' : lambda v,_ : int(float(v)) if v!='' else 0,
        'date' : lambda v,s : datetime.strptime(v,s['datefmt']),
        }
    def trim(self,val,triml): return self.trim(val.replace(triml[0],''),triml[1:]) \
        if triml else val
    def __init__(self,row,xos):
        for k,v in xos.fields.items():
            rawval = row[v['col']-1]
            trimval = self.trim(rawval,v['trim'])
            typfn = self.typd[v['typ']]
            conval = typfn(trimval,v)
            self.__dict__.update({k:conval})

class xls:
    def rows(self,flnm,sheet=0): return [
        [c.value for c in r] for r in xlrd.open_workbook(flnm).sheets()[sheet].get_rows()
        ]

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
    def __init__(self,flnm,specfile,sheet=0):
        spec = json.load(open(specfile))
        self.g = { **self._globals, **spec.get('globals',{}) }
        self.fields = { k:{ **self.g, **v } for k,v in spec['fields'].items() }
        strtrow = self.g['strtrow'] - 1
        endpatcol = self.g['endpatcol'] - 1
        endpat = self.g['endpat']
        ftyp = globals()[Path(flnm).suffix[1:]]
        self.objs = []
        for i,r in enumerate(ftyp().rows(flnm)):
            if i < strtrow : continue
            if r[endpatcol] == endpat: break
            self.objs = self.objs + [ XlsObj(r,self) ]

