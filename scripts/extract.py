"""Read the supplied BIFF8/WPS expense workbook; never writes the source.
Python standard library only. The .xls reader is deliberately bounded, not a general Excel engine.
"""
import argparse, collections, hashlib, io, json, re, struct, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

def u16(b, p=0): return struct.unpack_from('<H', b, p)[0]
def u32(b, p=0): return struct.unpack_from('<I', b, p)[0]
def column(n):
    s = ''; n += 1
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s
def address(r, c): return column(c) + str(r + 1)
def json_save(p, x): Path(p).write_text(json.dumps(x, ensure_ascii=False, indent=2), encoding='utf-8')

class Compound:
    def __init__(self, path):
        self.data = Path(path).read_bytes()
        if self.data[:8] != bytes.fromhex('d0cf11e0a1b11ae1'): raise ValueError('Expected OLE/BIFF .xls')
        self.ss = 1 << u16(self.data, 30)
        dif = list(struct.unpack_from('<109I', self.data, 76)); n = u32(self.data, 68)
        for _ in range(u32(self.data, 72)):
            d = self.sector(n); dif += list(struct.unpack_from('<%dI' % (self.ss//4-1), d)); n = u32(d, self.ss-4)
        self.fat = []
        for n in dif:
            if n < 0xfffffffa: self.fat += list(struct.unpack('<%dI' % (self.ss//4), self.sector(n)))
        d = self.chain(u32(self.data, 48)); self.entries = {}
        for i in range(0, len(d), 128):
            e = d[i:i+128]; l = u16(e, 64)
            if l: self.entries[e[:l-2].decode('utf-16le')] = (e[66], u32(e,116), struct.unpack_from('<Q',e,120)[0])
        root = self.entries['Root Entry']; self.mini = self.chain(root[1])[:root[2]]
        mf = self.chain(u32(self.data, 60)); self.minifat = list(struct.unpack('<%dI' % (len(mf)//4), mf))
    def sector(self, n): return self.data[(n+1)*self.ss:(n+2)*self.ss]
    def chain(self, n):
        out = []; seen = set()
        while n < 0xfffffffa:
            if n in seen or n >= len(self.fat): raise ValueError('Corrupt CFB chain')
            seen.add(n); out.append(self.sector(n)); n = self.fat[n]
        return b''.join(out)
    def stream(self, name):
        _, n, size = self.entries[name]
        if size >= u32(self.data,56): return self.chain(n)[:size]
        out = []; seen = set()
        while n < 0xfffffffa:
            if n in seen or n >= len(self.minifat): raise ValueError('Corrupt mini chain')
            seen.add(n); out.append(self.mini[n*64:(n+1)*64]); n = self.minifat[n]
        return b''.join(out)[:size]

class Segments:
    def __init__(self, chunks): self.cs=chunks; self.i=0; self.p=0
    def raw(self, n):
        out=b''
        while n:
            if self.p==len(self.cs[self.i]): self.i+=1; self.p=0
            take=min(n,len(self.cs[self.i])-self.p);out+=self.cs[self.i][self.p:self.p+take];self.p+=take;n-=take
        return out
    def num(self,n): return int.from_bytes(self.raw(n),'little')
    def string(self):
        n=self.num(2);flags=self.num(1);rich=self.num(2) if flags&8 else 0;ext=self.num(4) if flags&4 else 0;wide=flags&1;out=''
        while n:
            if self.p==len(self.cs[self.i]):self.i+=1;self.p=0;wide=self.num(1)&1
            take=min(n,(len(self.cs[self.i])-self.p)//(2 if wide else 1))
            if take == 0: raise ValueError('Invalid SST character split')
            out+=self.raw(take*(2 if wide else 1)).decode('utf-16le' if wide else 'latin1');n-=take
        self.raw(rich*4+ext);return out

def formula_decode(b, a):
    # Only forms actually verified in this fixture. Unknown tokens are exposed, never flattened silently.
    if b[:1] == b'\x01': return '=ROW()-2'  # validated against SHRFMLA below
    if b==b'\x1e\x00\x00':return '=0'
    if b==bytes.fromhex('420008001e020004'):return '=ROW()-2'
    if b[:5] == bytes.fromhex('2301000000') and b[5] == 0x17:
        n=b[6];name=b[8:8+n*2].decode('utf-16le');return '=_xlfn.DISPIMG("'+name+'",1)'
    if len(b)==13 and b[0] in (0x25,0x45,0x65) and b[-4:]==bytes.fromhex('42010400'):
        r1,r2,c1,c2=struct.unpack_from('<HHHH',b,1);return '=SUM('+address(r1,c1&255)+':'+address(r2,c2&255)+')'
    if len(b)==16 and b[:3]==bytes.fromhex('1e0900') and b[-4:]==bytes.fromhex('42025801'):
        r1,r2,c1,c2=struct.unpack_from('<HHHH',b,4);return '=SUBTOTAL(9,'+address(r1,c1&255)+':'+address(r2,c2&255)+')'
    raise ValueError('Unsupported formula '+a+': '+b.hex())

def read_xls(path):
    c=Compound(path);w=c.stream('Workbook' if 'Workbook' in c.entries else 'Book');records=[];pos=0
    while pos+4<=len(w):
        t,l=struct.unpack_from('<HH',w,pos)
        if pos+4+l>len(w):raise ValueError('Truncated BIFF record')
        records.append((pos,t,w[pos+4:pos+4+l]));pos+=4+l
    if any(t==0x2f for _,t,_ in records):raise ValueError('Encrypted workbook unsupported')
    strings=[];fonts=[];xfs=[];formats={};sheets=[]
    for i,(p,t,d) in enumerate(records):
        if t==0xfc:
            chunks=[d];j=i+1
            while j<len(records) and records[j][1]==0x3c:chunks.append(records[j][2]);j+=1
            r=Segments(chunks);r.num(4);n=r.num(4);strings=[r.string() for _ in range(n)]
        elif t==0x31:
            if len(fonts)==4:fonts.append(None)
            fonts.append({'size':u16(d)/20,'bold':u16(d,6)>=700,'italic':bool(u16(d,2)&2),'colorIndex':u16(d,4),'name':d[16:16+d[14]*(2 if d[15]&1 else 1)].decode('utf-16le' if d[15]&1 else 'latin1')})
        elif t==0xe0:xfs.append({'font':u16(d),'format':u16(d,2),'align':d[6]&7,'wrap':bool(d[6]&8),'valign':(d[6]>>4)&7,'hex':d.hex()})
        elif t==0x41e:formats[u16(d)]=Segments([d[2:]]).string()
        elif t==0x85:sheets.append({'name':d[8:8+d[6]*(2 if d[7]&1 else 1)].decode('utf-16le' if d[7]&1 else 'latin1'),'offset':u32(d),'cells':{},'styles':{},'formulas':{},'rows':{},'cols':[],'merges':[],'objects':0})
        elif t==0x4bc:
            if d[10:] != bytes.fromhex('420008001e020004'):raise ValueError('Unknown shared formula')
    byoffset={s['offset']:s for s in sheets};s=None
    def rk(n):
        v=(struct.unpack('<i',struct.pack('<I',n))[0]>>2) if n&2 else struct.unpack('<d',struct.pack('<II',0,n&0xfffffffc))[0]
        return v/100 if n&1 else v
    for p,t,d in records:
        if p in byoffset:s=byoffset[p]
        if s is None:continue
        if t in (0xfd,0x203,0x27e,0x6,0x205,0x201,0x204):
            a=address(u16(d),u16(d,2));s['styles'][a]=u16(d,4)
            if t==0x201:continue
            if t==0xfd:v=strings[u32(d,6)]
            elif t==0x204:v=Segments([d[6:]]).string()
            elif t==0x203:v=struct.unpack_from('<d',d,6)[0]
            elif t==0x27e:v=rk(u32(d,6))
            elif t==0x205:
                if d[7]:raise ValueError('Existing cell error '+s['name']+'!'+a)
                v=bool(d[6])
            else:
                v=struct.unpack_from('<d',d,6)[0] if d[12:14]!=b'\xff\xff' else None
                s['formulas'][a]=formula_decode(d[22:22+u16(d,20)],a)
            s['cells'][a]=v
        elif t==0xbd:
            row,col=u16(d),u16(d,2)
            for off in range(4,len(d)-2,6):
                a=address(row,col);s['cells'][a]=rk(u32(d,off+2));s['styles'][a]=u16(d,off);col+=1
        elif t==0xbe:
            row,col=u16(d),u16(d,2)
            for off in range(4,len(d)-2,2):s['styles'][address(row,col)]=u16(d,off);col+=1
        elif t==0x208:s['rows'][u16(d)+1]={'height':(u16(d,6)&0x7fff)/20,'hidden':bool(u16(d,12)&0x20)}
        elif t==0x7d:s['cols'].append({'first':u16(d),'last':u16(d,2),'width':u16(d,4)/256,'hidden':bool(u16(d,8)&1)})
        elif t==0xe5:
            for k in range(u16(d)):
                r1,r2,c1,c2=struct.unpack_from('<HHHH',d,2+8*k);s['merges'].append(address(r1,c1)+':'+address(r2,c2))
        elif t==0x5d:s['objects']+=1
    return {'sheets':sheets,'fonts':fonts,'xfs':xfs,'formats':formats}, c

def extract_images(c, out):
    if 'ETCellImageData' not in c.entries:raise ValueError('No WPS cell images in this XLS; provide image folder')
    data=c.stream('ETCellImageData');z=zipfile.ZipFile(io.BytesIO(data));ns={'xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    rels={a.attrib['Id']:'xl/'+a.attrib['Target'] for a in ET.fromstring(z.read('xl/_rels/cellImages.xml.rels'))};byhash={};ids={}
    for pic in ET.fromstring(z.read('xl/cellImages.xml')).findall('.//xdr:pic',ns):
        p=pic.find('xdr:nvPicPr/xdr:cNvPr',ns).attrib;b=pic.find('xdr:blipFill/a:blip',ns).attrib
        media=rels[b['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']];raw=z.read(media);sha=hashlib.sha256(raw).hexdigest();key=sha[:20]
        item=byhash.setdefault(sha,{'key':key,'sha256':sha,'names':[],'ids':[],'path':'images/'+key+Path(media).suffix})
        if p.get('descr','') not in item['names']:item['names'].append(p.get('descr',''))
        item['ids'].append(p['name']);ids[p['name']]=key
        dest=out/item['path'];dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(raw)
        from PIL import Image
        with Image.open(io.BytesIO(raw)) as im:item['width'],item['height']=im.size
    (out/'wps_cell_images.zip').write_bytes(data)
    # No sheet names, original coordinates, original captions or category hints in this corpus.
    return sorted(byhash.values(),key=lambda v:v['key']),ids

def run(source, out):
    source=Path(source).resolve();out=Path(out).resolve();out.mkdir(parents=True,exist_ok=True)
    wb,c=read_xls(source);imgs,ids=extract_images(c,out)
    raw=next(s for s in wb['sheets'] if s['name']=='未报');rows=[]
    for a,v in raw['cells'].items():
        if re.fullmatch(r'C\d+',a) and int(a[1:])>=3 and isinstance(v,str):
            n=int(a[1:]);amount=raw['cells'].get('D'+str(n));date=raw['cells'].get('B'+str(n))
            if not isinstance(amount,(int,float)):continue
            rows.append({'key':'未报!'+str(n),'source_row':n,'date':date,'project':v,'amount':amount,'note':raw['cells'].get('E'+str(n),'') or ''})
    rows.sort(key=lambda r:r['source_row'])
    json_save(out/'raw_rows.json',rows);json_save(out/'image_corpus.json',imgs)
    json_save(out/'template.json',wb);json_save(out/'evaluation_only_image_ids.json',ids)
    json_save(out/'provenance.json',{'source':str(source),'sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'records':len(rows),'unique_images':len(imgs),'image_definitions':len(ids)})
    return {'rows':len(rows),'images':len(imgs),'sheets':len(wb['sheets'])}

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('out');a=p.parse_args();print(json.dumps(run(a.source,a.out)))
