"""Integration regression: simulated monthly row counts; never modifies the supplied workbook."""
import argparse,collections,copy,json,struct,sys
from pathlib import Path
from extract import read_xls,json_save,run as extract_run
from adaptive import table,split,serialize,assemble,value_cell,xf,move_cell
from build_xls import build,numeric_formula
from match import run as match_run,cents
from verify_xls import verify

def fixture(source,rows,output):
    wb,c=read_xls(source);parts=split(c.stream('Workbook'),wb);dr,dc=table(parts['未报'])
    src=next(s for s in wb['sheets'] if s['name']=='未报');original_count=len([a for a in src['cells'] if a.startswith('C') and a[1:].isdigit() and int(a[1:])>=3 and isinstance(src['cells'].get('D'+a[1:]),(int,float))])
    rr={0:dr[0],1:dr[1]};cc={(r,col):v for (r,col),v in dc.items() if r<2}
    for i,row in enumerate(rows):
        r=i+2;donor=min(r,original_count+1);rr[r]=dr[donor]
        for col,value in enumerate([i+1,row['date'],row['project'],row['amount'],row['note']]):
            cc[r,col]=value_cell(r,col,xf(dc[donor,col]),value)
    # Source retains its one blank row before the total.
    n=len(rows);rr[n+2]=dr[original_count+2];rr[n+3]=dr[original_count+3]
    for (r,col),v in dc.items():
        if r==original_count+3:cc[n+3,col]=move_cell(v,n+3)
    token=b'\x45'+struct.pack('<HHHH',2,n+1,3,3)+bytes.fromhex('42010400') if n else b'\x1e\x00\x00'
    cc[n+3,3]=numeric_formula(n+3,3,xf(dc[original_count+3,3]),token,sum(cents(r['amount']) for r in rows)/100)
    output.write_bytes(assemble(source,wb,parts,{'未报':serialize(parts['未报'],rr,cc)},{'未报':n+3}))

def run(source,baseline,work,deps):
    baseline=Path(baseline);work=Path(work);work.mkdir(parents=True,exist_ok=True)
    raw=json.loads((baseline/'raw_rows.json').read_text(encoding='utf-8'));rules=Path(__file__).resolve().parents[1]/'assets/category-rules.json'
    cases={'empty':[],'one':raw[:1],'smaller':raw[:17],'larger':raw*3};reports={}
    for name,data in cases.items():
        folder=work/name;folder.mkdir(parents=True,exist_ok=True);sourcecopy=folder/'simulated-input.xls';output=folder/'simulated-output.xls'
        fixture(source,data,sourcecopy);extract_run(sourcecopy,folder);match_run(folder,rules);build(sourcecopy,folder,rules,output)
        reports[name]=verify(sourcecopy,output,folder,deps);assert reports[name]['records_reconciled']==len(data)
        print(name,json.dumps(reports[name],ensure_ascii=False),flush=True)
    json_save(work/'resize-test-results.json',reports)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('baseline');p.add_argument('work');p.add_argument('deps');a=p.parse_args();run(a.source,a.baseline,a.work,a.deps)
