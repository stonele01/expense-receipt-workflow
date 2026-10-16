"""Template-preserving BIFF8 content patch for the tested WPS expense layout.
Reuses every formatting/layout record and every non-Workbook OLE stream byte-for-byte.
Unknown layouts / missing target cells fail closed. No Office automation or macros are run.
"""
import argparse,bisect,datetime,hashlib,json,re,struct
from pathlib import Path
from extract import Compound,read_xls,u16,u32,address,json_save

FREE=0xffffffff;END=0xfffffffe;FATSECT=0xfffffffd;DIFSECT=0xfffffffc
def record(t,d):return struct.pack('<HH',t,len(d))+d
def unicode_string(s):return struct.pack('<HB',len(s),1)+s.encode('utf-16le')
def ref(sheet,row,col):return b'\x5a'+struct.pack('<HHH',sheet,row-1,col)
def numeric_formula(r,c,xf,tokens,value):
    return record(6,struct.pack('<HHH',r,c,xf)+struct.pack('<d',float(value))+struct.pack('<HIH',3,0,len(tokens))+tokens)
def text_formula(r,c,xf,tokens,value):
    return record(6,struct.pack('<HHH',r,c,xf)+bytes.fromhex('000000000000ffff')+struct.pack('<HIH',3,0,len(tokens))+tokens)+record(0x207,unicode_string(str(value)))
def blank(r,c,xf):return record(0x201,struct.pack('<HHH',r,c,xf))
def number(v):return format(float(v),'.12g')
def caption(row):
    d=datetime.datetime(1899,12,30)+datetime.timedelta(days=row['date'])
    return number(row['amount'])+(row['note'] or row['project'])+f'{d.month}/{d.day}'
def caption_tokens(sheet,row,use_note):
    # Dn & En (or Cn when note absent) & TEXT(Bn,"m/d"). '&' is CONCAT's equivalent.
    return ref(sheet,row,3)+ref(sheet,row,4 if use_note else 2)+b'\x08'+ref(sheet,row,1)+b'\x17\x03\x00m/d'+b'\x41\x30\x00'+b'\x08'
def make_cfb(original,workbook,stream_overrides=None):
    """Repack original CFB directory tree and streams, retaining directory metadata."""
    c=Compound(original);ss=c.ss
    if ss!=512 or u16(c.data,26)!=3:raise ValueError('Only version-3 CFB is supported')
    directory=bytearray(c.chain(u32(c.data,48)));chunks=[];fat=[]
    def alloc(data):
        start=len(chunks);count=(len(data)+ss-1)//ss
        if not count:return END,0
        for i in range(count):chunks.append(data[i*ss:(i+1)*ss].ljust(ss,b'\0'));fat.append(start+i+1 if i+1<count else END)
        return start,count
    mini=bytearray();minifat=[];root_offset=None
    for offset in range(0,len(directory),128):
        e=directory[offset:offset+128];l=u16(e,64)
        if not l:continue
        name=e[:l-2].decode('utf-16le');kind=e[66]
        if kind==5:root_offset=offset;continue
        if kind!=2:continue
        raw=workbook if name in ('Workbook','Book') else (stream_overrides or {}).get(name,c.stream(name))
        if len(raw)<4096:
            start=len(minifat);count=(len(raw)+63)//64
            if not count:start=END
            for i in range(count):mini.extend(raw[i*64:(i+1)*64].ljust(64,b'\0'));minifat.append(start+i+1 if i+1<count else END)
        else:start,_=alloc(raw)
        struct.pack_into('<IQ',directory,offset+116,start,len(raw))
    root_start,_=alloc(mini);struct.pack_into('<IQ',directory,root_offset+116,root_start,len(mini))
    mfdata=struct.pack('<%dI'%len(minifat),*minifat) if minifat else b''
    if mfdata:mfdata=mfdata.ljust(((len(mfdata)+ss-1)//ss)*ss,b'\xff')
    mfstart,mfcount=alloc(mfdata);dirstart,_=alloc(directory)
    n=len(chunks);fc=dc=0
    while True:
        nf=(n+fc+dc+127)//128;nd=max(0,(nf-109+126)//127)
        if (nf,nd)==(fc,dc):break
        fc,dc=nf,nd
    fids=list(range(n,n+fc));dids=list(range(n+fc,n+fc+dc));fat+=([FATSECT]*fc)+([DIFSECT]*dc)
    fat+=([FREE]*(fc*128-len(fat)))
    chunks.extend(struct.pack('<128I',*fat[i*128:(i+1)*128]) for i in range(fc))
    rest=fids[109:]
    for i in range(dc):
        part=rest[i*127:(i+1)*127];part+=[FREE]*(127-len(part));part.append(dids[i+1] if i+1<dc else END);chunks.append(struct.pack('<128I',*part))
    header=bytearray(c.data[:512]);struct.pack_into('<I',header,40,0);struct.pack_into('<I',header,44,fc);struct.pack_into('<I',header,48,dirstart)
    struct.pack_into('<IIII',header,60,mfstart,mfcount,dids[0] if dids else END,dc)
    hf=fids[:109]+[FREE]*max(0,109-len(fids));struct.pack_into('<109I',header,76,*hf)
    return bytes(header)+b''.join(chunks)

def build(source,folder,rules_path,output):
    source=Path(source).resolve();folder=Path(folder).resolve();output=Path(output).resolve()
    if source==output:raise ValueError('Output must be a separate copy')
    template,c=read_xls(source);w=c.stream('Workbook' if 'Workbook' in c.entries else 'Book')
    load=lambda name:json.loads((folder/name).read_text(encoding='utf-8'))
    if hashlib.sha256(source.read_bytes()).hexdigest()!=load('provenance.json')['sha256']:raise ValueError('Source workbook changed after extraction; extract and review again')
    matches=load('matches.json');corpus={i['key']:i for i in load('image_corpus.json')};rules=json.loads(Path(rules_path).read_text(encoding='utf-8'))
    if any(not r['category'] for r in matches):raise ValueError('Resolve unclassified records first')
    names=[s['name'] for s in template['sheets']];byname={s['name']:s for s in template['sheets']}
    if any(sum(r['category']==cfg['detail'] for r in matches)!=sum(bool(re.fullmatch(r'C\d+',a)) and int(a[1:])>=3 for a in byname[cfg['detail']]['cells']) for cfg in rules['sheets']) or any(not any('DISPIMG' in f for f in byname[cfg['images']]['formulas'].values()) for cfg in rules['sheets']):
        from adaptive import build_dynamic
        return build_dynamic(source,folder,rules,output)
    # Existing externsheet entries 0..2 refer to sheets 0..2. Preserve them and append remaining sheets.
    externs=[(0,i,i) for i in range(len(names))];updates={};placements=[]
    blocked={'unique_amount_candidate','missing_receipt','amount_conflict','ambiguous','unmatched','unclassified'}
    for cfg in rules['sheets']:
        selected=sorted((r for r in matches if r['category']==cfg['detail']),key=lambda r:r['source_row'])
        detail=byname[cfg['detail']];pics=byname[cfg['images']]
        available=sorted(int(a[1:]) for a in detail['cells'] if re.fullmatch(r'C\d+',a) and int(a[1:])>=3)
        imagecells=sorted((a for a,f in pics['formulas'].items() if 'DISPIMG' in f),key=lambda a:(int(re.search(r'\d+',a)[0]),a))
        if len(selected)!=len(available) or len(selected)!=len(imagecells):raise ValueError('Template capacity differs for '+cfg['detail']+'. Extend/review template in its native editor before using this preserving patch.')
        for i,row in enumerate(selected):
            n=available[i];sheet_idx=names.index(cfg['detail'])
            for col,value in enumerate([i+1,row['date'],row['project'],row['amount'],row['note']]):
                a=address(n-1,col)
                if a not in detail['styles']:raise ValueError('Missing styled detail cell '+a)
                tokens=bytes.fromhex('420008001e020004') if col==0 else ref(0,row['source_row'],col)
                updates[(cfg['detail'],a)]={'tokens':tokens,'value':value,'blank':value==''}
            a=imagecells[i];r=int(re.search(r'\d+',a)[0]);cc=re.match(r'[A-Z]+',a)[0];caption_cell=cc+str(r+1)
            if caption_cell not in pics['styles']:raise ValueError('Caption cell missing in template')
            key=row['image_key'] if row['status'] not in blocked else None
            if key:
                image=corpus[key];image_id=image['ids'][0]
                tokens=bytes.fromhex('2301000000')+b'\x17'+struct.pack('<BB',len(image_id),1)+image_id.encode('utf-16le')+bytes.fromhex('1e01004203ff00')
                updates[(cfg['images'],a)]={'tokens':tokens,'value':'=DISPIMG("'+image_id+'",1)'}
            else:updates[(cfg['images'],a)]={'blank':True,'value':None}
            updates[(cfg['images'],caption_cell)]={'tokens':caption_tokens(sheet_idx,n,bool(row['note'])),'value':caption(row)}
            placements.append({'row_key':row['key'],'sheet':cfg['images'],'cell':a,'caption':caption_cell,'detail_sheet':cfg['detail'],'detail_row':n,'status':row['status'],'image_key':key,'candidate_key':row['image_key'],'caption_value':caption(row)})
    records=[];p=0;s=None;offsets={s['offset']:s['name'] for s in template['sheets']}
    while p+4<=len(w):
        t,l=struct.unpack_from('<HH',w,p);d=w[p+4:p+4+l]
        if p in offsets:s=offsets[p]
        records.append((p,t,d,s));p+=4+l
    produced=[];mapping={};touched=set();skip_string=False;newpos=0
    def cell_replace(r,co,xf,old):
        key=(s,address(r,co))
        if key not in updates:return old,False
        u=updates[key];touched.add(key)
        if u.get('blank'):return blank(r,co,xf),True
        if isinstance(u['value'],str):return text_formula(r,co,xf,u['tokens'],u['value']),True
        return numeric_formula(r,co,xf,u['tokens'],u['value']),True
    for pos,t,d,s in records:
        mapping[pos]=newpos;raw=record(t,d)
        if t==0x207 and skip_string:raw=b'';skip_string=False
        elif t==0x17 and s is None:
            old=[struct.unpack_from('<HHH',d,2+i*6) for i in range(u16(d))]
            if old!=externs[:len(old)]:raise ValueError('Unsupported external references')
            raw=record(t,struct.pack('<H',len(externs))+b''.join(struct.pack('<HHH',*x) for x in externs))
        elif t in (0xfd,0x203,0x27e,0x6,0x201,0x205):
            raw,changed=cell_replace(u16(d),u16(d,2),u16(d,4),raw)
            skip_string=bool(changed and t==6 and d[12:14]==b'\xff\xff' and d[6]==0)
        elif t in (0xbd,0xbe):
            rr,cc=u16(d),u16(d,2);parts=[];step=6 if t==0xbd else 2
            for off in range(4,len(d)-2,step):
                xf=u16(d,off);old=record(0x27e,struct.pack('<HHH',rr,cc,xf)+d[off+2:off+6]) if t==0xbd else blank(rr,cc,xf)
                part,_=cell_replace(rr,cc,xf,old);parts.append(part);cc+=1
            # Splitting MULRK/MULBLANK preserves cell XFs and order and gives stable row-block offsets.
            raw=b''.join(parts)
        produced.append([pos,t,d,raw]);newpos+=len(raw)
    mapping[len(w)]=newpos
    starts=sorted(mapping)
    def translate(offset):
        start=starts[bisect.bisect_right(starts,offset)-1]
        return mapping[start]+offset-start
    if touched!=set(updates):raise ValueError('Some target cells absent: '+str(set(updates)-touched))
    # Repair all BIFF absolute/relative stream indexes after content-only record size changes.
    bsheet=0
    for pos,t,d,raw in produced:
        new=bytearray(d)
        if t==0x85:
            struct.pack_into('<I',new,0,mapping[u32(d)]);bsheet+=1
        elif t==0xff:
            for off in range(2,len(d),8):struct.pack_into('<I',new,off,translate(u32(d,off)))
        elif t==0x20b:
            for off in range(16,len(d),4):struct.pack_into('<I',new,off,mapping[u32(d,off)])
        elif t==0xd7:
            firstrow=pos-u32(d);struct.pack_into('<I',new,0,mapping[pos]-mapping[firstrow]);base=firstrow+4+u16(w,firstrow+2)
            for off in range(4,len(d),2):
                target=base+u16(d,off);delta=mapping[target]-mapping[base]
                if delta>65535:raise ValueError('Row block too large')
                struct.pack_into('<H',new,off,delta);base=target
        else:continue
        for item in produced:
            if item[0]==pos:item[3]=record(t,new);break
    newworkbook=b''.join(x[3] for x in produced);result=make_cfb(source,newworkbook)
    output.parent.mkdir(parents=True,exist_ok=True);output.write_bytes(result)
    # Independent verifiers must still check native formatting, OLE images, cells, captions and indexes.
    json_save(folder/'placements.json',placements)
    json_save(folder/'native-build.json',{'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'output':str(output),'changed_cells':len(updates),'inserted':sum(bool(p['image_key']) for p in placements),'review':sum(not p['image_key'] for p in placements),'format_strategy':'preserve original formatting/layout records and XFs; update only cell contents and stream indexes'})
    return {'output':str(output),'inserted':sum(bool(p['image_key']) for p in placements),'review':sum(not p['image_key'] for p in placements)}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('folder');p.add_argument('rules');p.add_argument('output');a=p.parse_args();print(json.dumps(build(a.source,a.folder,a.rules,a.output),ensure_ascii=False))
