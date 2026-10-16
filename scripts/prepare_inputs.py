"""Import an unreported XLS and images using the layout bundled with this skill."""
import argparse,datetime,hashlib,io,json,struct,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from PIL import Image
from extract import read_xls,address,json_save,run as extract_run
from adaptive import table,split,serialize,assemble,value_cell,xf,move_cell
from build_xls import numeric_formula
from match import cents

NS={'etc':'http://www.wps.cn/officeDocument/2017/etCustomData','xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
for prefix,uri in NS.items():ET.register_namespace(prefix,uri)
REL='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
DEFAULT_TEMPLATE=Path(__file__).resolve().parents[1]/'assets'/'layout-template.xls'
def image_package(paths):
    root=ET.Element('{'+NS['etc']+'}cellImages');rels=ET.Element('Relationships',xmlns=REL);types=ET.Element('Types',xmlns=CT);media={};manifest=[]
    for ext,mime in [('jpeg','image/jpeg'),('png','image/png'),('rels','application/vnd.openxmlformats-package.relationships+xml'),('xml','application/xml')]:ET.SubElement(types,'Default',Extension=ext,ContentType=mime)
    ET.SubElement(types,'Override',PartName='/xl/cellImages.xml',ContentType='application/vnd.wps-officedocument.cellimage+xml')
    def sub(p,ns,tag,**attrs):return ET.SubElement(p,'{'+NS[ns]+'}'+tag,attrs)
    for i,path in enumerate(paths,1):
        raw=path.read_bytes();sha=hashlib.sha256(raw).hexdigest()
        with Image.open(io.BytesIO(raw)) as im:
            if im.format not in ('JPEG','PNG'):raise ValueError('Only JPEG/PNG are supported: '+path.name)
            width,height=im.size;ext='png' if im.format=='PNG' else 'jpeg'
        media_name='media/'+sha+'.'+ext;media['xl/'+media_name]=raw;rid='rId'+str(i);imageid='ID_'+hashlib.sha256(raw+path.name.encode('utf-8')).hexdigest()[:32].upper()
        item=sub(root,'etc','cellImage');pic=sub(item,'xdr','pic');nv=sub(pic,'xdr','nvPicPr');sub(nv,'xdr','cNvPr',id=str(i+1),name=imageid,descr=path.stem);sub(nv,'xdr','cNvPicPr')
        blip=sub(pic,'xdr','blipFill');b=sub(blip,'a','blip');b.set('{'+NS['r']+'}embed',rid);sub(sub(blip,'a','stretch'),'a','fillRect')
        sp=sub(pic,'xdr','spPr');transform=sub(sp,'a','xfrm');sub(transform,'a','off',x='0',y='0');sub(transform,'a','ext',cx=str(width*9525),cy=str(height*9525));sub(sub(sp,'a','prstGeom',prst='rect'),'a','avLst')
        ET.SubElement(rels,'Relationship',Id=rid,Type=NS['r']+'/image',Target=media_name)
        manifest.append({'file':path.name,'sha256':sha})
    rootrels=ET.Element('Relationships',xmlns=REL);ET.SubElement(rootrels,'Relationship',Id='rId1',Type='http://www.wps.cn/officeDocument/2020/cellImage',Target='xl/cellImages.xml')
    parts={'[Content_Types].xml':ET.tostring(types,encoding='utf-8',xml_declaration=True),'_rels/.rels':ET.tostring(rootrels,encoding='utf-8',xml_declaration=True),'xl/cellImages.xml':ET.tostring(root,encoding='utf-8',xml_declaration=True),'xl/_rels/cellImages.xml.rels':ET.tostring(rels,encoding='utf-8',xml_declaration=True),**media}
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w') as z:
        for name,data in sorted(parts.items()):
            info=zipfile.ZipInfo(name,(1980,1,1,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED;z.writestr(info,data)
    return out.getvalue(),manifest

def run(source,images,folder,template=None):
    source=Path(source).resolve();images=Path(images).resolve();template=Path(template).resolve() if template is not None else DEFAULT_TEMPLATE;folder=Path(folder).resolve()
    if not template.is_file():raise FileNotFoundError('Bundled layout missing; restore assets/layout-template.xls from the complete skill package: '+str(template))
    folder.mkdir(parents=True,exist_ok=True)
    if not images.is_dir():raise ValueError('Image folder not found')
    sw,_=read_xls(source);raw=next(s for s in sw['sheets'] if s['name']=='未报');entries=[]
    for a,v in raw['cells'].items():
        if a.startswith('C') and a[1:].isdigit() and int(a[1:])>=3 and isinstance(v,str):
            n=int(a[1:]);amount=raw['cells'].get('D'+str(n));date=raw['cells'].get('B'+str(n))
            if isinstance(amount,(int,float)):entries.append({'source_row':n,'date':date,'project':v,'amount':amount,'note':raw['cells'].get('E'+str(n),'') or ''})
    entries.sort(key=lambda r:r['source_row']);tw,c=read_xls(template);parts=split(c.stream('Workbook'),tw);dr,dc=table(parts['未报']);rows={r:d for r,d in dr.items() if r<2};cells={(r,col):v for (r,col),v in dc.items() if r<2}
    for a in ('A1','A2','B2','C2','D2','E2'):
        r=int(a[1:])-1;col=ord(a[0])-65
        if a in raw['cells']:cells[r,col]=value_cell(r,col,xf(dc[r,col]),raw['cells'][a])
    # Preserve the actual source row IDs, including any gaps in the input.
    for i,item in enumerate(entries):
        r=item['source_row']-1;rows[r]=dr[2]
        for col,v in enumerate([i+1,item['date'],item['project'],item['amount'],item['note']]):cells[r,col]=value_cell(r,col,xf(dc[2,col]),v)
    last=max([1]+[r['source_row']-1 for r in entries]);total=last+2
    for r in range(2,total+1):rows.setdefault(r,dr[2])
    total_template=max(r for (r,col),v in dc.items() if col==3)
    rows[total]=dr[total_template]
    for (r,col),v in dc.items():
        if r==total_template:cells[total,col]=move_cell(v,total)
    tokens=b'\x45'+struct.pack('<HHHH',2,last,3,3)+bytes.fromhex('42010400') if entries else b'\x1e\x00\x00'
    cells[total,3]=numeric_formula(total,3,xf(dc[total_template,3]),tokens,sum(cents(r['amount']) for r in entries)/100)
    paths=sorted(p for p in images.iterdir() if p.is_file() and p.suffix.lower() in ('.jpg','.jpeg','.png'));package,manifest=image_package(paths)
    prepared=folder/'prepared-source.xls'
    if prepared in (source,template):raise ValueError('Prepared workbook must not replace input/template')
    prepared.write_bytes(assemble(template,tw,parts,{'未报':serialize(parts['未报'],rows,cells)},{'未报':total},{'ETCellImageData':package}))
    result=extract_run(prepared,folder)
    json_save(folder/'input-bundle.json',{'unreported_file':str(source),'unreported_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'template_sha256':hashlib.sha256(template.read_bytes()).hexdigest(),'screenshots':manifest})
    return {'prepared_source':str(prepared),**result}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('images');p.add_argument('folder');p.add_argument('--template',help='Optional replacement layout; defaults to the skill bundled template');a=p.parse_args();print(json.dumps(run(a.source,a.images,a.folder,a.template),ensure_ascii=False))
