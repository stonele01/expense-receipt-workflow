import random,unittest
from match import build,tokens,cents
class MatcherTests(unittest.TestCase):
    rules={'prefixes':{'办公费：':'办公费'},'exact':{'办公费':['文件夹','白板笔','铅笔']}}
    def row(self,key,project,amount,note=''):
        return dict(key=key,project=project,amount=amount,note=note,date=46113,source_row=int(key))
    def image(self,key,name):return dict(key=key,names=[name],sha256=key)
    def test_amount_only_is_never_a_verified_match(self):
        r=build([self.row('1','文件夹',25.42)],[self.image('a','25.42支付截图')],self.rules)[0]
        self.assertEqual(r['status'],'unique_amount_candidate')
    def test_duplicate_names_remain_ambiguous(self):
        rs=build([self.row('1','文件夹',10)],[self.image('a','10文件夹'),self.image('b','10文件夹')],self.rules)
        self.assertEqual(rs[0]['status'],'ambiguous');self.assertIsNone(rs[0]['image_key'])
    def test_unknown_category_is_not_defaulted(self):
        r=build([self.row('1','未知用途',10)],[],self.rules)[0];self.assertEqual(r['status'],'unclassified')
    def test_no_fuzzy_amount_matching(self):
        r=build([self.row('1','文件夹',28.8)],[self.image('a','28.88文件夹')],self.rules)[0];self.assertIsNone(r['image_key'])
    def test_receipt_group_must_reconcile(self):
        rows=[self.row('1','白板笔',1),self.row('2','铅笔',1)];images=[self.image('a','2办公用品')]
        with self.assertRaises(ValueError):build(rows,images,self.rules,{'groups':[{'rows':['1','2'],'image_key':'a','receipt_amount':3,'reason':'test'}]})
        rs=build(rows,images,self.rules,{'groups':[{'rows':['1','2'],'image_key':'a','receipt_amount':2,'reason':'itemized receipt'}]});self.assertTrue(all(r['status']=='shared_receipt' for r in rs))
    def test_duplicate_image_needs_explicit_group(self):
        rows=[self.row('1','白板笔',1),self.row('2','铅笔',1)];images=[self.image('a','1支付')]
        d={'matches':{k:{'image_key':'a','status':'user_confirmed','reason':'test'} for k in ['1','2']}}
        with self.assertRaises(ValueError):build(rows,images,self.rules,d)
    def test_date_and_amount_are_distinct(self):
        self.assertIn(4624,tokens('46.24打车'));self.assertNotIn(417,tokens('张三4.17号打车36.31'));self.assertIn(3631,tokens('张三4.17号打车36.31'))
    def test_confirmation_changes_only_selected_row(self):
        rows=[self.row('1','白板笔',3),self.row('2','铅笔',4)];images=[self.image('a','3支付'),self.image('b','4支付')]
        d={'matches':{'1':{'image_key':'a','image_sha256':'a','status':'user_confirmed','reason':'用户确认用途'}}}
        rs=build(rows,images,self.rules,d);self.assertEqual(rs[0]['status'],'user_confirmed');self.assertEqual(rs[1]['status'],'unique_amount_candidate')
    def test_shuffled_input_does_not_change_distinct_matches(self):
        rows=[self.row('1','白板笔',3),self.row('2','文件夹',3)];images=[self.image('a','3文件夹'),self.image('b','3白板笔')]
        a={r['key']:r['image_key'] for r in build(rows,images,self.rules)};random.Random(8).shuffle(rows);random.Random(9).shuffle(images)
        self.assertEqual(a,{r['key']:r['image_key'] for r in build(rows,images,self.rules)})
if __name__=='__main__':unittest.main()
