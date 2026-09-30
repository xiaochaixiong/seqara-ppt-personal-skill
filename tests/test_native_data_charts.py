import sys
import unittest
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from native_pptx import package, TEMPLATE, ordered_slides, xml, shapes, relationships
from native_data_charts import add_data_chart, validate_data_chart
from quality import validate_chart


class NativeDataChartTests(unittest.TestCase):
    def setUp(self):
        self.facts={f'm{i}':{'value':v,'unit':'项','period':f'2026-{i:02d}'}
                    for i,v in enumerate([8,12,21,15],1)}
        self.template=package(TEMPLATE)
        self.slide=ordered_slides(self.template)[3]
        self.root=xml(self.template[self.slide])
        self.rels=relationships(self.template,self.slide)
        self.page={'slots':{str(i):{'action':'remove'} for i in range(8,46)}}
        self.spec={'type':'trend','categories':['一月','二月','三月','四月'],
                   'series':[{'name':'项目数','fact_ids':['m1','m2','m3','m4']}],
                   'bounds_pt':[45,135,870,340], 'replace_shapes':[str(i) for i in range(8,46)],
                   'reason':'Data chart replaces historical chart region',
                   'source_label':'来源：合成样例','interpretation':'三月最高'}

    def test_five_editable_chart_grammars_and_cache_workbook(self):
        for typ in ('trend','comparison','ranking','composition','target_progress'):
            with self.subTest(typ=typ):
                spec=dict(self.spec,type=typ)
                if typ in ('comparison','target_progress'):
                    spec['series']=[self.spec['series'][0],{'name':'目标','fact_ids':['m4']*4}]
                if typ=='ranking':
                    spec['series']=[{'name':'项目数','fact_ids':['m3','m4','m2','m1']}]
                    spec['categories']=['三月','四月','二月','一月']
                facts=self.facts if typ not in ('composition','ranking','target_progress') else {
                    k:dict(v,period='2026上半年') for k,v in self.facts.items()}
                validate_data_chart(spec,self.page,shapes(self.root),facts)
                parts=dict(self.template)
                root=xml(self.template[self.slide]);rels=relationships(self.template,self.slide)
                record=add_data_chart(parts,root,rels,'ppt/slides/personal-test.xml',1,spec,facts)
                self.assertEqual(record['type'],typ)
                audit=validate_chart(parts,record['chart_part'])
                self.assertEqual(len(audit['series']),len(spec['series']))
                self.assertEqual(audit['series'][0]['values'],
                                 [float(self.facts[f]['value']) for f in spec['series'][0]['fact_ids']])

    def test_claim_and_replacement_guards(self):
        bad=dict(self.spec,replace_shapes=['6'])
        with self.assertRaisesRegex(ValueError,'identity|title'):
            validate_data_chart(bad,self.page,shapes(self.root),self.facts)
        bad=dict(self.spec,series=[{'name':'项目数','fact_ids':['m1','m2','m3','missing']}])
        with self.assertRaisesRegex(ValueError,'missing'):
            validate_data_chart(bad,self.page,shapes(self.root),self.facts)
        bad=dict(self.spec,type='composition')
        separate=dict(self.facts)
        separate['m2']=dict(separate['m2'],unit='人')
        with self.assertRaisesRegex(ValueError,'unit'):
            validate_data_chart(bad,self.page,shapes(self.root),separate)
        bad=dict(self.spec,type='composition',categories=[str(i) for i in range(7)],
                 series=[{'name':'项目数','fact_ids':['m1']*7}])
        with self.assertRaisesRegex(ValueError,'six'):
            validate_data_chart(bad,self.page,shapes(self.root),self.facts)


if __name__=='__main__':unittest.main()
