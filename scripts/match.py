"""Match raw expense rows against an unordered image-name corpus.
The matcher never reads the template or its prior picture positions.
"""
import argparse, collections, datetime, json, re, unicodedata
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

def cents(v):return int((Decimal(str(v))*100).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
def load(p):return json.loads(Path(p).read_text(encoding='utf-8'))
def save(p,x):Path(p).write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf-8')
def norm(s):
    s=unicodedata.normalize('NFKC',s).lower()
    for a,b in [('网开','网卡'),('文件框','文件筐')]:s=s.replace(a,b)
    return re.sub(r'[^\u4e00-\u9fffa-z]','',s)
def tokens(s):
    s=re.sub(r'(?<!\d)(?:0?[1-9]|1[0-2])[./](?:0?[1-9]|[12]\d|3[01])(?=号|日|晚餐|奶茶|打车|团建|回家)','',s)
    s=re.sub(r'\d{1,2}月\d{1,2}(?:日|号)?','',s)
    return {cents(v) for v in re.findall(r'(?<![\d.])\d+(?:\.\d{1,2})?(?![\d.])',s)}
def semantic(a,b):
    a,b=norm(a),norm(b);common=[]
    for n in (4,3,2):
        for i in range(len(a)-n+1):
            v=a[i:i+n]
            if v in b and v not in ('支付','截图','费用','办公','办公室','加班','福利','其他','公室') and not any(v in x for x in common):common.append(v)
    return sum(len(v)**2 for v in common),common
def classify(row,rules):
    for prefix,cat in rules['prefixes'].items():
        if row['project'].startswith(prefix):return cat
    for cat,labels in rules['exact'].items():
        if row['project'] in labels:return cat
    return None
def build(rows,images,rules,decisions=None):
    decisions=decisions or {'matches':{},'groups':[]}
    result=[];used=collections.defaultdict(list);iby={i['key']:i for i in images}
    for row in rows:
        item=dict(row,category=classify(row,rules),status='unmatched',image_key=None,evidence=[],candidates=[])
        for im in images:
            names=' '.join(im['names']);sc,words=semantic(row['project']+' '+row['note'],names)
            exact=cents(row['amount']) in tokens(names)
            if exact or sc>=12:item['candidates'].append({'key':im['key'],'names':im['names'],'amount_exact':exact,'score':sc,'shared_words':words})
        item['candidates'].sort(key=lambda v:(v['amount_exact'],v['score'],v['key']),reverse=True)
        result.append(item)
    # Strong lexical matches first. A globally unique amount is a weaker, explicitly labelled candidate.
    for item in sorted(result,key=lambda r: r['candidates'][0]['score'] if r['candidates'] else -1,reverse=True):
        exact=[a for a in item['candidates'] if a['amount_exact'] and not used[a['key']]]
        if not exact:continue
        best=exact[0];runner=exact[1] if len(exact)>1 else None
        if runner and best['score']-runner['score']<4: item['status']='ambiguous';continue
        # Shared source signatures are allowed to remain ambiguous; do not distinguish by template order.
        competing=[r for r in result if r is not item and cents(r['amount'])==cents(item['amount']) and not r['image_key']]
        if best['score']==0 and competing:item['status']='ambiguous';continue
        item['image_key']=best['key'];item['status']='name_amount' if best['score']>=4 else 'unique_amount_candidate'
        item['evidence']=['图片名称金额相同']+best['shared_words'];used[best['key']].append(item['key'])
    # Reviewed decisions are separate evidence, never embedded in general matching rules.
    byrow={r['key']:r for r in result}
    for rowkey,d in decisions.get('matches',{}).items():
        if rowkey not in byrow or (d.get('image_key') and d['image_key'] not in iby):raise ValueError('Unknown reviewed key')
        if not d.get('reason'):raise ValueError('Reviewed decision needs evidence')
        if d['status'] not in {'equivalent_pair','visual_review','amount_conflict','missing_receipt','user_confirmed','user_confirmed_image_conflict'}:raise ValueError('Unsupported reviewed status')
        if d.get('image_sha256') and d['image_sha256']!=iby[d['image_key']]['sha256']:raise ValueError('Image changed since confirmation')
        r=byrow[rowkey];r['image_key']=d.get('image_key');r['status']=d['status'];r['evidence']=[d['reason']]
    approved_groups=[]
    for g in decisions.get('groups',[]):
        if not g.get('reason') or g['image_key'] not in iby:raise ValueError('Invalid shared receipt')
        selected=[byrow[k] for k in g['rows']]
        if sum(cents(r['amount']) for r in selected)!=cents(g['receipt_amount']):raise ValueError('Split receipt does not reconcile')
        approved_groups.append(set(g['rows']))
        for r in selected:r['image_key']=g['image_key'];r['status']='shared_receipt';r['evidence']=[g['reason']]
    usage=collections.defaultdict(list)
    for r in result:
        if r['image_key']:usage[r['image_key']].append(r['key'])
    for key,rs in usage.items():
        if len(rs)>1 and not any(set(rs)==g for g in approved_groups):raise ValueError('Unapproved duplicate image use: '+str(rs))
    for r in result:
        r['candidates']=r['candidates'][:6]
        if not r['category']:r['status']='unclassified'
    return result
def run(folder,rules_file,decisions_file=None):
    folder=Path(folder);rows=load(folder/'raw_rows.json');images=load(folder/'image_corpus.json');rules=load(rules_file)
    decisions=load(decisions_file) if decisions_file else None
    if decisions and decisions.get('source_sha256') and decisions['source_sha256']!=load(folder/'provenance.json')['sha256']:raise ValueError('Confirmation belongs to a different workbook')
    matches=build(rows,images,rules,decisions)
    save(folder/'matches.json',matches)
    stats={'records':len(matches),'amount':sum(cents(r['amount']) for r in matches)/100,'status':dict(collections.Counter(r['status'] for r in matches)),'categories':{}}
    for cat in set(r['category'] for r in matches):
        rs=[r for r in matches if r['category']==cat];stats['categories'][str(cat)]={'count':len(rs),'amount':sum(cents(r['amount']) for r in rs)/100}
    save(folder/'match_stats.json',stats);return stats
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('folder');p.add_argument('--rules',required=True);p.add_argument('--decisions');a=p.parse_args();print(json.dumps(run(a.folder,a.rules,a.decisions),ensure_ascii=False))
