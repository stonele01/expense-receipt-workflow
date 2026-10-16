"""Resize the known expense template by cloning native rows and cells.

BIFF row blocks, INDEX, DIMENSIONS, totals and filter ranges are regenerated.
Fonts/XFs, column widths, page setup, header drawings and WPS image streams stay native.
"""
import bisect,collections,datetime,hashlib,json,re,struct
from pathlib import Path
from extract import read_xls,u16,u32,address,json_save
from build_xls import record,blank,ref,numeric_formula,text_formula,caption,caption_tokens,make_cfb
from match import cents

CELL={0xfd,0x203,0x27e,6,0x201,0x205,0x204}
TABLE=CELL|{0xbd,0xbe,0x207,0x208,0xd7,0x4bc}
BLOCKED={'unique_amount_candidate','missing_receipt','amount_conflict','ambiguous','unmatched','unclassified'}

def records(w):
    pos=0;out=[]
    while pos+4<=len(w):
        t,l=struct.unpack_from('<HH',w,pos)
        if pos+4+l>len(w):raise ValueError('Truncated BIFF record')
        out.append((pos,t,w[pos+4:pos+4+l]));pos+=4+l
    if pos!=len(w):raise ValueError('Unparsed BIFF tail')
    return out

def split(w,template):
    bounds={s['offset']:s['name'] for s in template['sheets']};parts={None:[]};name=None
    for p,t,d in records(w):
        if p in bounds:name=bounds[p];parts[name]=[]
        parts[name].append((p,t,d))
    return parts

def table(rs):
    rows={};cells={};last=None
    for p,t,d in rs:
        if t==0x208:rows[u16(d)]=d
        elif t in CELL:
            last=(u16(d),u16(d,2));cells[last]=record(t,d)
        elif t==0x207:
            if last is None:raise ValueError('Orphan formula string')
            cells[last]+=record(t,d)
        elif t in (0xbd,0xbe):
            r,c=u16(d),u16(d,2);step=6 if t==0xbd else 2
            for off in range(4,len(d)-2,step):
                xf=u16(d,off)
                cells[r,c]=record(0x27e,struct.pack('<HHH',r,c,xf)+d[off+2:off+6]) if t==0xbd else blank(r,c,xf)
                c+=1
    return rows,cells

def move_cell(raw,r,c=None):
    data=bytearray(raw);struct.pack_into('<H',data,4,r)
    if c is not None:struct.pack_into('<H',data,6,c)
    return bytes(data)

def xf(raw):return u16(raw,8)
def value_cell(r,c,style,value,tokens=None):
    if value is None or value=='':return blank(r,c,style)
    if tokens is not None:
        return text_formula(r,c,style,tokens,value) if isinstance(value,str) else numeric_formula(r,c,style,tokens,value)
    if isinstance(value,str):
        # BIFF8 LABEL, avoiding any modification of the original SST and its references.
        encoded=value.encode('utf-16le')
        if len(encoded)>8000:raise ValueError('Cell text too long for bounded LABEL writer')
        return record(0x204,struct.pack('<HHHHB',r,c,style,len(encoded)//2,1)+encoded)
    return record(0x203,struct.pack('<HHHd',r,c,style,float(value)))

def serialize(rs,rows,cells):
    """Return a sheet with relative INDEX pointers (rebased on workbook assembly)."""
    if not rows or max(rows)>65535:raise ValueError('XLS worksheet row limit exceeded')
    first,last=min(rows),max(rows)+1
    loc=[i for i,(_,t,_) in enumerate(rs) if t in TABLE]
    if any(t not in TABLE for _,t,_ in rs[min(loc):max(loc)+1]):raise ValueError('Unsupported records interleaved with cell table')
    prefix=[(t,d) for _,t,d in rs[:min(loc)]];suffix=[(t,d) for _,t,d in rs[max(loc)+1:]]
    nblocks=(last-first+31)//32;cols=[c for r,c in cells];maxcol=max(cols)+1 if cols else 0
    for i,(t,d) in enumerate(prefix):
        if t==0x20b:prefix[i]=(t,struct.pack('<4I',0,first,last,0)+bytes(nblocks*4))
        elif t==0x200:prefix[i]=(t,struct.pack('<IIHHH',first,last,min(cols) if cols else 0,maxcol,0))
    out=bytearray(b''.join(record(t,d) for t,d in prefix));dbs=[]
    byrow=collections.defaultdict(list)
    for (r,c),raw in sorted(cells.items()):byrow[r].append(raw)
    for start in range(first,last,32):
        block=[r for r in range(start,min(start+32,last)) if r in rows]
        firstrow=len(out);firstlen=None
        for r in block:
            d=bytearray(rows[r]);struct.pack_into('<H',d,0,r);cs=[u16(raw,6) for raw in byrow[r]]
            struct.pack_into('<HH',d,2,min(cs) if cs else 0,max(cs)+1 if cs else 0)
            raw=record(0x208,d);out.extend(raw)
            if firstlen is None:firstlen=len(raw)
        offsets=[];base=firstrow+(firstlen or 0)
        for r in block:
            if byrow[r]:
                delta=len(out)-base
                if delta>65535:raise ValueError('BIFF row block offset exceeds 65535')
                offsets.append(delta);base=len(out);out.extend(b''.join(byrow[r]))
        dbs.append(len(out));out.extend(record(0xd7,struct.pack('<I',len(out)-firstrow if offsets else 0)+struct.pack('<%dH'%len(offsets),*offsets)))
    out.extend(b''.join(record(t,d) for t,d in suffix))
    parsed=records(out);idx=next(p for p,t,d in parsed if t==0x20b);defcol=next(p for p,t,d in parsed if t==0x55)
    struct.pack_into('<I',out,idx+4+12,defcol)
    for i,p in enumerate(dbs):struct.pack_into('<I',out,idx+4+16+i*4,p)
    return bytes(out)

def assemble(source,template,parts,replacements,new_ranges=None,stream_overrides=None):
    """Rebase all native absolute stream offsets; unchanged sheets remain byte-identical otherwise."""
    names=[s['name'] for s in template['sheets']];global_out=[];offsetmap={};newpos=0
    for p,t,d in parts[None]:
        offsetmap[p]=newpos
        if t==0x17:
            old=[struct.unpack_from('<HHH',d,2+i*6) for i in range(u16(d))]
            externs=[(0,i,i) for i in range(len(names))]
            if old!=externs[:len(old)]:raise ValueError('Unsupported external references')
            d=struct.pack('<H',len(externs))+b''.join(struct.pack('<HHH',*x) for x in externs)
        if t==0x18 and new_ranges:
            itab=u16(d,8);name=names[itab-1] if itab else None
            if name in new_ranges and u16(d)&0x20 and d[3]==1:
                # Built-in FilterDatabase or Print_Area with a simple local ptgArea3d.
                token_start=16
                if d[15] in (6,13) and d[token_start:token_start+1] in (b'\x3b',b'\x5b',b'\x7b') and u16(d,4)==11:
                    d=bytearray(d);struct.pack_into('<H',d,token_start+5,new_ranges[name]);d=bytes(d)
                elif d[15] not in (7,13):raise ValueError('Unsupported dynamic named range')
        raw=record(t,d);global_out.append([p,t,d,raw]);newpos+=len(raw)
    sheet_offsets={};chunks=[]
    for name in names:
        sheet_offsets[name]=newpos
        if name in replacements:raw=replacements[name];delta=newpos
        else:
            raw=b''.join(record(t,d) for _,t,d in parts[name]);delta=newpos-parts[name][0][0]
        data=bytearray(raw)
        parsed=records(raw);defcol=next(p for p,t,d in parsed if t==0x55)
        for p,t,d in parsed:
            if t==0x20b:
                struct.pack_into('<I',data,p+4+12,newpos+defcol)
                for off in range(16,len(d),4):
                    if u32(d,off):struct.pack_into('<I',data,p+4+off,u32(d,off)+delta)
        chunks.append(bytes(data));newpos+=len(data)
    for item in global_out:
        p,t,d,raw=item
        if t==0x85:
            oldoff=u32(d);name=next(s['name'] for s in template['sheets'] if s['offset']==oldoff)
            d=bytearray(d);struct.pack_into('<I',d,0,sheet_offsets[name]);item[3]=record(t,d)
        elif t==0xff:
            d=bytearray(d)
            starts=sorted(offsetmap)
            for off in range(2,len(d),8):
                old=u32(d,off);base=starts[bisect.bisect_right(starts,old)-1]
                struct.pack_into('<I',d,off,offsetmap[base]+old-base)
            item[3]=record(t,d)
    return make_cfb(source,b''.join(x[3] for x in global_out)+b''.join(chunks),stream_overrides)

def build_dynamic(source,folder,rules,output):
    source=Path(source);folder=Path(folder);output=Path(output)
    load=lambda n:json.loads((folder/n).read_text(encoding='utf-8'))
    template,c=read_xls(source);parts=split(c.stream('Workbook'),template);byname={s['name']:s for s in template['sheets']};names=list(byname)
    matches=load('matches.json');images={x['key']:x for x in load('image_corpus.json')};raw=load('raw_rows.json')
    if len({r['key'] for r in matches})!=len(raw) or {r['key'] for r in matches}!={r['key'] for r in raw}:raise ValueError('Source coverage mismatch')
    rawby={r['key']:r for r in raw}
    for r in matches:
        if any(r[k]!=rawby[r['key']][k] for k in ('source_row','date','project','amount','note')):raise ValueError('Matched data differs from raw input')
    months={(datetime.datetime(1899,12,30)+datetime.timedelta(days=r['date'])).month for r in raw}
    replacements={};layouts={};placements=[];ranges={}
    for cfg in rules['sheets']:
        name=cfg['detail'];picname=cfg['images'];selected=sorted((r for r in matches if r['category']==name),key=lambda r:r['source_row']);n=len(selected)
        detail=byname[name];pic=byname[picname];dr,dc=table(parts[name]);pr,pc=table(parts[picname])
        oldn=max(dr)-2
        if not oldn or max(dr)!=oldn+2:raise ValueError('Expected header, details and one total row: '+name)
        rows={0:dr[0],1:dr[1]};cells={(r,c):v for (r,c),v in dc.items() if r<2};rowmap={0:0,1:1};stylemap={address(r,c):address(r,c) for r,c in cells}
        for i,item in enumerate(selected):
            r=i+2;donor=min(r,oldn+1);rows[r]=dr[donor];rowmap[r]=donor
            for col,v in enumerate([i+1,item['date'],item['project'],item['amount'],item['note']]):
                templatecell=dc.get((donor,col),dc.get((2,col)))
                if templatecell is None:raise ValueError('Detail style missing')
                tokens=bytes.fromhex('420008001e020004') if col==0 else ref(0,item['source_row'],col)
                cells[r,col]=value_cell(r,col,xf(templatecell),v,tokens);stylemap[address(r,col)]=address(donor,col) if (donor,col) in dc else address(2,col)
        total=n+2;rows[total]=dr[oldn+2];rowmap[total]=oldn+2
        for (r,col),v in dc.items():
            if r==oldn+2:cells[total,col]=move_cell(v,total);stylemap[address(total,col)]=address(r,col)
        tokens=b'\x45'+struct.pack('<HHHH',2,n+1,3,3)+bytes.fromhex('42010400') if n else b'\x1e\x00\x00'
        cells[total,3]=numeric_formula(total,3,xf(dc[oldn+2,3]),tokens,sum(cents(r['amount']) for r in selected)/100)
        if len(months)==1 and isinstance(detail['cells'].get('A1'),str):
            title=re.sub(r'^\d+月份',str(next(iter(months)))+'月份',detail['cells']['A1'])
            cells[0,0]=value_cell(0,0,xf(dc[0,0]),title)
        replacements[name]=serialize(parts[name],rows,cells);ranges[name]=total
        layouts[name]={'kind':'detail','count':n,'row_donors':rowmap,'cell_donors':stylemap,'total_row':total+1}
        # Infer the actual original grid; recruitment/traffic use two columns, not five.
        slots=[(int(re.search(r'\d+',a)[0])-1,ord(a[0])-65) for a,f in pic['formulas'].items() if 'DISPIMG' in f]
        if slots:
            start=min(r for r,c in slots);left=min(c for r,c in slots);width=max(c for r,c in slots)-left+1
            oldpairs=(max(r for r,c in slots)-start)//2+1
            if set(slots)!={(start+2*(i//width),left+i%width) for i in range(oldn)}:raise ValueError('Irregular image grid: '+picname)
        else:
            start=cfg['start_row']-1;left=cfg['start_col']-1;width=cfg['columns']
            tall=[int(r)-1 for r,v in pic['rows'].items() if v['height']>100]
            if not tall or min(tall)!=start or tall!=list(range(start,max(tall)+1,2)):raise ValueError('Blank template image pattern missing')
            oldpairs=(max(tall)-start)//2+1
        pairs=max(1,(n+width-1)//width);end=start+2*pairs
        # A validated helper sequence in columns before the pictures is kept current.
        helpers=[]
        for col in range(left):
            seq=sorted((int(a[1:])-1,v) for a,v in pic['cells'].items() if a.startswith(chr(65+col)) and isinstance(v,(int,float)))
            if seq:
                expected=[start+3+2*i if i<oldpairs-1 else start+2*oldpairs+i-(oldpairs-1) for i in range(oldn)]
                if [v for _,v in seq]!=list(range(1,oldn+1)) or [r for r,_ in seq]!=expected:raise ValueError('Unknown helper data')
                positions=[start+3+2*i if i<pairs-1 else end+i-(pairs-1) for i in range(n)]
                helpers.append((col,positions))
        last_content=start+2*oldpairs+(max(0,oldn-(oldpairs-1)) if helpers else 0);padding=max(pr)+1-last_content
        extent=max([end]+[max(pos)+1 for _,pos in helpers if pos])+max(0,padding)
        rows={};cells={};rowmap={};stylemap={}
        for r in range(min(pr),extent):
            if start<=r<end:donor=start+2*min((r-start)//2,oldpairs-1)+(r-start)%2
            elif r in pr and r>=start+2*oldpairs:donor=r
            else:donor=start+1
            rows[r]=pr[donor];rowmap[r]=donor
        for i in range(pairs*width):
            r=start+2*(i//width);col=left+i%width
            for rr in (r,r+1):
                donor=rr if (rr,col) in pc and rr<start+2*oldpairs else start+(rr-start)%2
                rawcell=pc.get((donor,col))
                if rawcell is None:raise ValueError('Image grid style missing')
                cells[rr,col]=blank(rr,col,xf(rawcell));stylemap[address(rr,col)]=address(donor,col)
            if i>=n:continue
            item=selected[i];key=item['image_key'] if item['status'] not in BLOCKED else None
            if key:
                imageid=images[key]['ids'][0];tokens=bytes.fromhex('2301000000')+b'\x17'+struct.pack('<BB',len(imageid),1)+imageid.encode('utf-16le')+bytes.fromhex('1e01004203ff00')
                cells[r,col]=text_formula(r,col,xf(cells[r,col]),tokens,'=DISPIMG("'+imageid+'",1)')
            cells[r+1,col]=text_formula(r+1,col,xf(cells[r+1,col]),caption_tokens(names.index(name),i+3,bool(item['note'])),caption(item))
            placements.append({'row_key':item['key'],'sheet':picname,'cell':address(r,col),'caption':address(r+1,col),'detail_sheet':name,'detail_row':i+3,'status':item['status'],'image_key':key,'candidate_key':item['image_key'],'caption_value':caption(item)})
        for col,positions in helpers:
            header=start+1
            cells[header,col]=pc[header,col];stylemap[address(header,col)]=address(header,col)
            for i,r in enumerate(positions):
                donor=r if (r,col) in pc else start+3
                cells[r,col]=value_cell(r,col,xf(pc[donor,col]),i+1);stylemap[address(r,col)]=address(donor,col)
        replacements[picname]=serialize(parts[picname],rows,cells)
        layouts[picname]={'kind':'images','count':n,'start_row':start+1,'start_col':left+1,'columns':width,'pairs':pairs,'row_donors':rowmap,'cell_donors':stylemap,'helpers':helpers}
    output.parent.mkdir(parents=True,exist_ok=True);output.write_bytes(assemble(source,template,parts,replacements,ranges))
    json_save(folder/'placements.json',placements);json_save(folder/'layout-plan.json',layouts)
    report={'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'output':str(output),'inserted':sum(bool(p['image_key']) for p in placements),'review':sum(not p['image_key'] for p in placements),'format_strategy':'adaptive native rows: template cell/row styles, fixed original column widths; regenerate sums, grids and BIFF indexes'}
    json_save(folder/'native-build.json',report);return report
