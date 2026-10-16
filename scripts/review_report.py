"""Create a self-contained local candidate review page. No uploads, network calls or account access."""
import argparse,base64,datetime,html,io,json
from pathlib import Path
from PIL import Image

LABEL={'unique_amount_candidate':'仅金额相同，商品或用途不能确认','amount_conflict':'截图金额与账表金额不同','missing_receipt':'未找到可确认的凭证','ambiguous':'存在多个候选','unmatched':'未找到匹配','unclassified':'分类未确认'}
def run(folder,output,decisions=None):
    folder=Path(folder);load=lambda n:json.loads((folder/n).read_text(encoding='utf-8'))
    matches=load('matches.json');images={x['key']:x for x in load('image_corpus.json')};prov=load('provenance.json')
    pending=[r for r in matches if r['status'] in LABEL];base=json.loads(Path(decisions).read_text(encoding='utf-8')) if decisions else {'matches':{},'groups':[]}
    base['source_sha256']=prov['sha256'];manifest={};sections=[];thumbs={}
    def thumb(key):
        if key not in thumbs:
            with Image.open(folder/images[key]['path']) as im:
                im=im.convert('RGB');im.thumbnail((620,1450));buf=io.BytesIO();im.save(buf,format='JPEG',quality=85);thumbs[key]='data:image/jpeg;base64,'+base64.b64encode(buf.getvalue()).decode()
        return thumbs[key]
    esc=lambda s:html.escape(str(s),quote=True)
    for r in pending:
        keys=[]
        if r['status']!='missing_receipt':
            if r['image_key']:keys.append(r['image_key'])
            keys += [a['key'] for a in r['candidates'] if a['amount_exact'] and a['key'] not in keys]
        keys=keys[:3];manifest[r['key']]={'status':r['status'],'keys':keys};cards=[]
        for key in keys:
            im=images[key]
            cards.append(f'<label class="candidate"><input type="radio" name="{esc(r["key"])}" value="{key}"><span>确认使用这张图</span><p>{esc(" / ".join(im["names"]))}</p><img loading="lazy" src="{thumb(key)}" alt="候选支付凭证"></label>')
        d=datetime.datetime(1899,12,30)+datetime.timedelta(days=r['date']);notes='；'.join(r['evidence'])
        sections.append(f'<section data-status="{r["status"]}"><h2>{esc(r["key"])} · {r["amount"]:.2f} 元 · {esc(r["project"])}</h2><p>{d:%Y-%m-%d}　{esc(r["category"])}　{esc(r["note"])}</p><p class="reason">{esc(LABEL[r["status"]])}</p><p>{esc(notes)}</p><div class="candidates">'+''.join(cards)+f'</div><label class="keep"><input type="radio" name="{esc(r["key"])}" value="" checked>暂不确认，继续留空</label></section>')
    payload=json.dumps({'base':base,'rows':manifest,'images':{k:{'sha256':images[k]['sha256']} for r in manifest.values() for k in r['keys']}},ensure_ascii=False).replace('</','<'+chr(92)+'/')
    doc='''<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>费用截图待核对清单</title><style>
    *{box-sizing:border-box}body{margin:0;background:#f7f7f4;color:#242424;font:16px/1.65 "Microsoft YaHei",sans-serif}header,main{max-width:1120px;margin:auto;padding:28px}h1{font-size:30px;margin:0}h2{font-size:19px;margin:0 0 8px}p{margin:8px 0}.toolbar{position:sticky;top:0;background:#fff;border-bottom:1px solid #ddd;padding:14px 28px;z-index:2;display:flex;gap:20px;align-items:center;flex-wrap:wrap}button,select{font:inherit;padding:9px 15px;border:1px solid #999;border-radius:5px;background:white;cursor:pointer}button{background:#175448;color:#fff;border:0}section{background:white;border:1px solid #ddd;margin:0 0 24px;padding:24px}.reason{color:#a23c22;font-weight:bold}.candidates{display:flex;gap:20px;overflow-x:auto}.candidate{display:block;flex:0 0 min(100%,430px);border:1px solid #ddd;padding:12px;cursor:pointer}.candidate:has(input:checked){outline:3px solid #175448}.candidate img{display:block;max-width:100%;height:auto;margin-top:10px}.candidate p{font-size:14px}.keep{display:block;margin-top:16px}input{width:18px;height:18px;vertical-align:middle}.hint{color:#555;font-size:14px}@media(max-width:600px){header,main{padding:18px}section{padding:16px}}
    </style><header><h1>费用截图待核对清单</h1><p>161 笔费用已分类。117 个位置已填图，44 个位置暂留空。</p><p>本页列出候选，不表示已经识别出商品。仅在你能确认用途时勾选；信息不足就继续留空。</p><p class="hint">所有图片均来自原工作簿，页面可离线使用。点击“导出确认结果”只生成确认文件，不会直接修改工作簿。金额冲突即使确认图片，也不自动修改原金额。</p></header><div class="toolbar"><select id="filter"><option value="all">全部待核对</option><option value="unique_amount_candidate">仅金额相同</option><option value="amount_conflict">金额冲突</option><option value="missing_receipt">缺少凭证</option></select><span id="count">本次确认 0 笔</span><button id="export">导出确认结果</button></div><main>'''+''.join(sections)+'''</main><script id="payload" type="application/json">'''+payload+'''</script><script>
    const data=JSON.parse(document.getElementById('payload').textContent);
    function collect(){const chosen={};for(const input of document.querySelectorAll('input[type=radio]:checked'))if(input.value)chosen[input.name]=input.value;return chosen;}
    function makeDecisions(chosen){const result=JSON.parse(JSON.stringify(data.base));result.matches=result.matches||{};for(const [row,key] of Object.entries(chosen)){if(!data.rows[row]||!data.rows[row].keys.includes(key))throw Error('Invalid selection');const conflict=data.rows[row].status==='amount_conflict';result.matches[row]={image_key:key,image_sha256:data.images[key].sha256,status:conflict?'user_confirmed_image_conflict':'user_confirmed',reason:conflict?'用户在核对清单确认图片对应此费用；金额差异仍保留，未修改账表金额。':'用户在核对清单确认图片对应此费用。'};}return result;}
    document.addEventListener('change',()=>{document.getElementById('count').textContent='本次确认 '+Object.keys(collect()).length+' 笔';});
    document.getElementById('filter').addEventListener('change',e=>{for(const section of document.querySelectorAll('section'))section.hidden=e.target.value!=='all'&&section.dataset.status!==e.target.value;});
    document.getElementById('export').addEventListener('click',()=>{const chosen=collect();if(!Object.keys(chosen).length){alert('还没有确认任何图片。可以保持留空，确认后再导出。');return;}const blob=new Blob([JSON.stringify(makeDecisions(chosen),null,2)],{type:'application/json;charset=utf-8'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='截图确认结果.json';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);});
    </script></html>'''
    # Counts are data-derived for subsequent runs.
    inserted=sum(r['status'] not in LABEL and r['image_key'] is not None for r in matches)
    doc=doc.replace('161 笔费用已分类。117 个位置已填图，44 个位置暂留空。',f'{len(matches)} 笔费用已分类。{inserted} 个位置已填图，{len(pending)} 个位置暂留空。')
    Path(output).write_text(doc,encoding='utf-8');return {'review_rows':len(pending),'candidate_images':len(thumbs)}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('folder');p.add_argument('output');p.add_argument('--decisions');a=p.parse_args();print(json.dumps(run(a.folder,a.output,a.decisions)))
