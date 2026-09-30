import sys
import unittest
from pathlib import Path

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from pipeline import validate_binding


class NumericSignTests(unittest.TestCase):
    def test_negative_source_cannot_authorize_positive_slide_claim(self):
        refs={'source:p001':'实际变化为 -5 次'}
        sources={'source':{'role':'primary','sha256':'a'}}
        good={'kind':'source','action':'replace','text':'变化 -5 次','sources':['source:p001']}
        validate_binding(good,'',refs,{},sources,'template')
        bad=dict(good,text='变化 5 次')
        with self.assertRaisesRegex(ValueError,'Unbound visible number'):
            validate_binding(bad,'',refs,{},sources,'template')


if __name__=='__main__':unittest.main()
