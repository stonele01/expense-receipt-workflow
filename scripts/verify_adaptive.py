"""Independent xlrd/olefile checks for variable-size native template output."""
import collections,hashlib,json,re,struct,sys
from pathlib import Path
from extract import address,u16,u32,json_save
from match import cents

def verify(source,output,folder,deps):
    sys.path.insert(0,str(Path(deps).resolve()));import xlrd,olefile
    from xlrd.formula import decompile_formula,FMLA_TYPE_CELL
    from verify_xls import records
    folder=Path(folder);load=lambda n:json.loads((folder/n).read_text(encoding='utf-8'))
    before=xlrd.open_workbook(source,formatting_info=True);after=xlrd.open_workbook(output,formatting_info=True)
    assert before.sheet_names()==after.sheet_names()
    oldole=olefile.OleFileIO(source);newole=olefile.OleFileIO(output,raise_defects=olefile.DEFECT_INCORRECT)
    assert oldole.listdir()==newole.listdir() and not newole.parsing_issues
    for e in oldole.listdir():
        if e[-1] not in ('Workbook','Book'):assert oldole.openstream(e).read()==newole.openstream(e).read()
    old=records(oldole.openstream('Workbook').read());new=records(newole.openstream('Workbook').read())
    plans=load('layout-plan.json');matches=load('matches.json');places=load('placements.json');imgs={i['key']:i for i in load('image_corpus.json')}
    assert {r['key'] for r in matches}=={r['key'] for r in load('raw_rows.json')}
    # Validate every saved style against a real template donor, and all column / merge / page settings.
    for a,b in zip(before.sheets(),after.sheets()):
        assert a.merged_cells==b.merged_cells
        assert {c:vars(v) for c,v in a.colinfo_map.items()}=={c:vars(v) for c,v in b.colinfo_map.items()}
        if a.name not in plans:
            assert [a.row_values(r) for r in range(a.nrows)]==[b.row_values(r) for r in range(b.nrows)]
            assert {r:vars(v) for r,v in a.rowinfo_map.items()}=={r:vars(v) for r,v in b.rowinfo_map.items()}
            for r in range(a.nrows):
                for c in range(a.ncols):assert a.cell_xf_index(r,c)==b.cell_xf_index(r,c)
            continue
        plan=plans[a.name]
        for r,donor in plan['row_donors'].items():assert vars(b.rowinfo_map[int(r)])==vars(a.rowinfo_map[donor]),(a.name,r,'row format')
        for addr,donor in plan['cell_donors'].items():
            rr=int(re.search(r'\d+',addr)[0])-1;cc=ord(addr[0])-65;dr=int(re.search(r'\d+',donor)[0])-1;dc=ord(donor[0])-65
            assert b.cell_xf_index(rr,cc)==a.cell_xf_index(dr,dc),(a.name,addr,'cell format')
        if plan['kind']=='detail':
            selected=sorted((r for r in matches if r['category']==a.name),key=lambda r:r['source_row']);n=len(selected)
            assert b.nrows==n+3 and plan['count']==n and plan['total_row']==n+3
            assert b.row_values(1)==a.row_values(1)
            for i,row in enumerate(selected):
                assert b.row_values(i+2)[:5]==[i+1,row['date'],row['project'],row['amount'],row['note']]
                raw=before.sheet_by_name('未报');rr=row['source_row']-1
                assert raw.row_values(rr)[1:5]==[row['date'],row['project'],row['amount'],row['note']]
                donor=min(i+2,a.nrows-2)
                assert vars(b.rowinfo_map[i+2])==vars(a.rowinfo_map[donor])
                for c in range(5):assert b.cell_xf_index(i+2,c)==a.cell_xf_index(donor,c)
            assert cents(b.cell_value(n+2,3))==sum(cents(r['amount']) for r in selected)
            assert b.cell_xf_index(n+2,3)==a.cell_xf_index(a.nrows-1,3)
        else:
            width=plan['columns'];start=plan['start_row']-1;left=plan['start_col']-1;n=plan['count']
            assert plan['pairs']==max(1,(n+width-1)//width)
            for r in range(start,start+2*plan['pairs']):
                assert b.rowinfo_map[r].height==a.rowinfo_map[start+(r-start)%2].height
            for i in range(n,plan['pairs']*width):
                r=start+2*(i//width);c=left+i%width
                assert b.cell_value(r,c)==b.cell_value(r+1,c)==''
            for col,positions in plan['helpers']:assert [b.cell_value(i,col) for i in positions]==list(range(1,n+1))
    def groups(rs):
        names=before.sheet_names();bounds={u32(d):names[i] for i,(_,t,d) in enumerate(x for x in rs if x[1]==0x85)};out=collections.defaultdict(list);name=None
        for p,t,d in rs:
            name=bounds.get(p,name);out[name].append((p,t,d))
        return out
    og,ng=groups(old),groups(new)
    celltypes={0x20b,0x200,0x208,0xd7,0xfd,0x203,0x27e,6,0x201,0x205,0x204,0xbd,0xbe,0x207,0x4bc}
    for name in before.sheet_names():
        assert [(t,d) for _,t,d in og[name] if t not in celltypes]==[(t,d) for _,t,d in ng[name] if t not in celltypes],(name,'page / column / drawing records')
    assert [(t,d) for _,t,d in og[None] if t not in (0x85,0x17,0x18,0xff)]==[(t,d) for _,t,d in ng[None] if t not in (0x85,0x17,0x18,0xff)],'Fonts/XFs changed'
    starts={p:(t,d) for p,t,d in new};formula_count=0;formulas={}
    for name,rs in ng.items():
        for p,t,d in rs:
            if t==0x85:assert starts[u32(d)][0]==0x809
            elif t==0x20b:
                if u32(d,12):assert starts[u32(d,12)][0]==0x55
                for off in range(16,len(d),4):assert starts[u32(d,off)][0]==0xd7
            elif t==0xd7 and u32(d):
                first=p-u32(d);assert starts[first][0]==0x208;base=first+4+len(starts[first][1])
                for off in range(4,len(d),2):base+=u16(d,off);assert starts[base][0] in (0xfd,0x203,0x27e,6,0x201,0x205,0xbd,0xbe,0x204)
            elif t==6:
                tok=d[22:22+u16(d,20)]
                if tok[:1]==b'\x01':continue
                formula=decompile_formula(after,tok,len(tok),FMLA_TYPE_CELL,browx=u16(d),bcolx=u16(d,2));assert formula and '#REF!' not in formula and 'ERR' not in formula
                formulas[name,address(u16(d),u16(d,2))]=formula;formula_count+=1
    for name,p in plans.items():
        if p['kind']=='detail':
            n=p['count'];f=formulas[name,'D'+str(n+3)]
            assert f.replace('$','')==(f'SUM(D3:D{n+2})' if n else '0.0'),(name,f)
            selected=sorted((r for r in matches if r['category']==name),key=lambda r:r['source_row'])
            for i,row in enumerate(selected):
                for c,field in [(1,'date'),(2,'project'),(3,'amount'),(4,'note')]:
                    if row[field]=='':continue
                    expected='未报!$'+chr(65+c)+'$'+str(row['source_row'])
                    assert formulas[name,address(i+2,c)]==expected,(name,i,c,'wrong source formula')
    byrow={r['key']:r for r in matches};assert len(places)==len(byrow) and {p['row_key'] for p in places}==set(byrow)
    for p in places:
        row=byrow[p['row_key']];s=after.sheet_by_name(p['sheet']);r=int(re.search(r'\d+',p['cell'])[0])-1;c=ord(p['cell'][0])-65
        assert s.cell_value(r+1,c)==p['caption_value']
        dn=p['detail_row'];detail=p['detail_sheet'];textcol='E' if row['note'] else 'C'
        expected=f'{detail}!$D${dn}&{detail}!${textcol}${dn}&TEXT({detail}!$B${dn},"m/d")'
        assert formulas[p['sheet'],address(r+1,c)]==expected,(p['sheet'],p['caption'],'wrong caption formula')
        if p['image_key']:assert imgs[p['image_key']]['ids'][0] in s.cell_value(r,c)
        else:assert s.cell_value(r,c)==''
        grid=plans[p['sheet']];i=p['detail_row']-3
        assert (r,c)==(grid['start_row']-1+2*(i//grid['columns']),grid['start_col']-1+i%grid['columns'])
    actual_pictures=sum(isinstance(v,str) and v.startswith('=DISPIMG(') for name in plans if plans[name]['kind']=='images' for r in range(after.sheet_by_name(name).nrows) for v in after.sheet_by_name(name).row_values(r))
    assert actual_pictures==sum(bool(p['image_key']) for p in places),'Stale or missing pictures'
    view=load('template.json')
    for s in view['sheets']:
        b=after.sheet_by_name(s['name']);s['cells']={address(r,c):b.cell_value(r,c) for r in range(b.nrows) for c in range(b.ncols) if b.cell_type(r,c) not in (0,6)};s['formulas']={}
        s['styles']={address(r,c):b.cell_xf_index(r,c) for r in range(b.nrows) for c in range(b.ncols)}
        s['rows']={r+1:{'height':v.height/20,'hidden':bool(v.hidden)} for r,v in b.rowinfo_map.items()}
    json_save(folder/'native-view.json',view)
    report={'source_unchanged':hashlib.sha256(Path(source).read_bytes()).hexdigest()==load('provenance.json')['sha256'],'adaptive_layout':True,'records_reconciled':len(matches),'category_counts':{n:p['count'] for n,p in plans.items() if p['kind']=='detail'},'inserted':sum(bool(p['image_key']) for p in places),'review':sum(not p['image_key'] for p in places),'native_styles_and_columns_preserved':True,'unrelated_sheets_unchanged':True,'formulas_and_indexes_valid':True,'decoded_formulas':formula_count,'native_wps_ui_recalculation':'not tested'}
    assert report['source_unchanged'];json_save(folder/'verification.json',report);return report
