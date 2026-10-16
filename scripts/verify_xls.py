"""Independent value/style/container validation using xlrd and olefile, installed in a local validation dir."""
import argparse,collections,hashlib,json,re,struct,sys
from pathlib import Path
from extract import Compound,read_xls,u16,u32,address,json_save
from match import cents

def records(data):
    p=0;out=[]
    while p+4<=len(data):
        t,l=struct.unpack_from('<HH',data,p);assert p+4+l<=len(data);out.append((p,t,data[p+4:p+4+l]));p+=4+l
    assert p==len(data);return out
def verify(source,output,folder,deps):
    metadata=json.loads((Path(folder)/'native-build.json').read_text(encoding='utf-8'))
    if metadata.get('format_strategy','').startswith('adaptive'):
        from verify_adaptive import verify as verify_dynamic
        return verify_dynamic(source,output,folder,deps)
    sys.path.insert(0,str(Path(deps).resolve()));import xlrd,olefile
    from xlrd.formula import decompile_formula,FMLA_TYPE_CELL
    folder=Path(folder);load=lambda n:json.loads((folder/n).read_text(encoding='utf-8'))
    src=xlrd.open_workbook(source,formatting_info=True);dst=xlrd.open_workbook(output,formatting_info=True)
    assert src.sheet_names()==dst.sheet_names()
    co=olefile.OleFileIO(source);cn=olefile.OleFileIO(output,raise_defects=olefile.DEFECT_INCORRECT)
    assert not cn.parsing_issues and co.listdir()==cn.listdir()
    identical_streams=[]
    for entry in co.listdir():
        if entry[-1] not in ('Workbook','Book'):
            assert co.openstream(entry).read()==cn.openstream(entry).read(),entry
            identical_streams.append(entry[-1])
    old=co.openstream('Workbook').read();new=cn.openstream('Workbook').read();oldrs=records(old);newrs=records(new)
    variable={0x85,0x17,0xff,0x20b,0xd7,0xfd,0x203,0x27e,0x6,0x201,0x205,0xbd,0xbe,0x207}
    assert [(t,d) for _,t,d in oldrs if t not in variable]==[(t,d) for _,t,d in newrs if t not in variable], 'Format/layout or unrelated records changed'
    # Check numeric row-block indexes, not just sequential parsing.
    starts={p:(t,d) for p,t,d in newrs}
    for p,t,d in newrs:
        if t==0x85:assert starts[u32(d)][0]==0x809
        elif t==0x20b:
            for off in range(16,len(d),4):assert starts[u32(d,off)][0]==0xd7
        elif t==0xd7:
            first=p-u32(d);assert starts[first][0]==0x208;base=first+4+len(starts[first][1])
            for off in range(4,len(d),2):
                base+=u16(d,off);assert base in starts
    for before,after in zip(src.sheets(),dst.sheets()):
        assert (before.nrows,before.ncols)==(after.nrows,after.ncols)
        for r in range(before.nrows):
            for c in range(before.ncols):assert before.cell_xf_index(r,c)==after.cell_xf_index(r,c),(before.name,r,c)
        assert before.merged_cells==after.merged_cells
        assert {r:vars(v) for r,v in before.rowinfo_map.items()}=={r:vars(v) for r,v in after.rowinfo_map.items()}
        assert {r:vars(v) for r,v in before.colinfo_map.items()}=={r:vars(v) for r,v in after.colinfo_map.items()}
    for name in ['未报','总','Sheet3']:
        a,b=src.sheet_by_name(name),dst.sheet_by_name(name)
        assert [a.row_values(r) for r in range(a.nrows)]==[b.row_values(r) for r in range(b.nrows)]
    matches={r['key']:r for r in load('matches.json')};places=load('placements.json');imgs={i['key']:i for i in load('image_corpus.json')}
    seen=[];sums=collections.defaultdict(int);counts=collections.Counter()
    for p in places:
        r=matches[p['row_key']];seen.append(r['key']);detail=dst.sheet_by_name(p['detail_sheet']);n=p['detail_row']-1
        assert detail.row_values(n)[:5]==[n-1,r['date'],r['project'],r['amount'],r['note']]
        sums[p['detail_sheet']]+=cents(r['amount']);counts[p['detail_sheet']]+=1
        s=dst.sheet_by_name(p['sheet']);rr=int(re.search(r'\d+',p['cell'])[0])-1;cc=ord(p['cell'][0])-65
        assert s.cell_value(rr+1,cc)==p['caption_value']
        if p['image_key']:
            actual=s.cell_value(rr,cc);assert 'DISPIMG' in actual and imgs[p['image_key']]['ids'][0] in actual
        else:assert s.cell_value(rr,cc)==''
    assert len(seen)==len(set(seen))==len(matches)
    for name,total in sums.items():assert cents(dst.sheet_by_name(name).cell_value(counts[name]+2,3))==total
    # Independently decompile newly written BIFF formulas (cross-sheet reference and TEXT/concat).
    sheet=None;formula_examples=[];formula_count=0
    bounds={u32(d):src.sheet_names()[i] for i,(_,_,d) in enumerate([x for x in newrs if x[1]==0x85])}
    for pos,t,d in newrs:
        if pos in bounds:sheet=bounds[pos]
        if t!=6:continue
        tokens=d[22:22+u16(d,20)];r,c=u16(d),u16(d,2)
        if tokens[:1]==b'\x01':continue
        text=decompile_formula(dst,tokens,len(tokens),FMLA_TYPE_CELL,browx=r,bcolx=c)
        assert text and 'ERR' not in text and '#REF!' not in text,(sheet,r,c,text)
        formula_count+=1
        if len(formula_examples)<12 and sheet in ('办公费','办公费支付截图'):formula_examples.append({'sheet':sheet,'cell':address(r,c),'formula':text})
    # Actual output values and XFs for rendering; not the planned values.
    view=load('template.json')
    for s in view['sheets']:
        native=dst.sheet_by_name(s['name']);s['cells']={address(r,c):native.cell_value(r,c) for r in range(native.nrows) for c in range(native.ncols) if native.cell_type(r,c) not in (0,6)};s['formulas']={}
    json_save(folder/'native-view.json',view)
    report={'source_unchanged':hashlib.sha256(Path(source).read_bytes()).hexdigest()==load('provenance.json')['sha256'],'sheet_names_and_order_identical':True,'all_cell_styles_identical':True,'all_row_heights_and_column_widths_identical':True,'all_merged_ranges_identical':True,'format_layout_records_byte_identical':True,'unchanged_ole_streams':identical_streams,'ole_parsing_issues':[],'biff_stream_indexes_valid':True,'records_reconciled':len(seen),'totals':{k:v/100 for k,v in sums.items()},'inserted':sum(bool(p['image_key']) for p in places),'review':sum(not p['image_key'] for p in places),'decoded_nonshared_formulas':formula_count,'formula_examples':formula_examples,'native_wps_ui_recalculation':'not tested'}
    assert report['source_unchanged'];json_save(folder/'verification.json',report);return report
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');p.add_argument('folder');p.add_argument('deps');a=p.parse_args();print(json.dumps(verify(a.source,a.output,a.folder,a.deps),ensure_ascii=False))
