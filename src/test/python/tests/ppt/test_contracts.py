import copy
import unittest
from wps_skills.core.action_session import ContractValidationError
from wps_skills.ppt.contracts import PPT_TARGET_CONTRACT_SET as contracts


class PptContractsTests(unittest.TestCase):
    def test_examples_validate_and_all_edits_require_observations(self):
        for c in contracts.contracts:
            for example in c.examples: contracts.validate_params(c.name,example['params'])
            if c.risk=='write' and c.name!='save':
                example=dict(c.examples[0]['params']);del example['expectedToken']
                with self.assertRaises(ContractValidationError): contracts.validate_params(c.name,example)

    def test_paths_cannot_be_relative_wrong_format_or_device_names(self):
        for path in (r'C:\work\slides.pptx',r'\\server\share\slides.pptx'):
            contracts.validate_params('openPresentation',{'path':path})
        for path in ('slides.pptx','/tmp/slides.pptx',r'C:\slides.pptm',r'C:\NUL.pptx',r'C:\bad.\slides.pptx','C:\\bad\n.pptx'):
            with self.subTest(path=path),self.assertRaises(ContractValidationError): contracts.validate_params('openPresentation',{'path':path})

    def test_mutation_values_are_bounded_and_do_not_coerce(self):
        base={'slideId':256,'shapeId':2,'expectedToken':'observed'}
        for text in ('a\r\nb','a\x00b','\ud800','😀'*5001):
            with self.subTest(text=repr(text[:20])),self.assertRaises(ContractValidationError):
                contracts.validate_params('setShapeText',dict(base,text=text))
        contracts.validate_params('setShapeText',dict(base,text='中文\n😀\ttext'))
        for patch in ({'format':{}},{'format':{'size':float('nan')}},{'format':{'bold':1}},{'format':{'color':-1}}):
            with self.assertRaises(ContractValidationError): contracts.validate_params('formatText',dict(base,**patch))

    def test_shape_readback_must_prove_identity_text_font_and_geometry(self):
        shape={'id':2,'name':'Text','type':17,'left':10,'top':20,'width':200,'height':80,'rotation':0,
               'text':'Updated','font':{'latinName':'Arial','eastAsianName':'宋体','size':32,'bold':True,'italic':False,'color':255}}
        snapshot={'slide':{'id':256,'index':1,'name':'Slide1','layout':12,'shapeCount':1},'shapes':[shape],'token':'after'}
        base={'slideId':256,'shapeId':2,'expectedToken':'before'}
        cases=[('setShapeText',{'text':'Updated'},'text','wrong'),
               ('formatText',{'format':{'bold':True}},'font',dict(shape['font'],bold=False)),
               ('setShapeGeometry',{'geometry':{'left':10}},'left',99)]
        for action,params,key,bad in cases:
            contracts.validate_result(action,snapshot,params=dict(base,**params))
            broken=copy.deepcopy(snapshot);broken['shapes'][0][key]=bad
            with self.assertRaises(ContractValidationError): contracts.validate_result(action,broken,params=dict(base,**params))
        broken=copy.deepcopy(snapshot);broken['slide']['id']=999
        with self.assertRaises(ContractValidationError): contracts.validate_result('getSlideInfo',broken,params={'slideId':256})

    def test_slide_move_and_delete_require_observed_results(self):
        slide={'id':256,'index':1,'name':'Slide1','layout':12,'shapeCount':0}
        result={'slides':[slide],'token':'after'}
        with self.assertRaises(ContractValidationError): contracts.validate_result('moveSlide',result,params={'slideId':256,'position':2,'expectedToken':'before'})
        with self.assertRaises(ContractValidationError): contracts.validate_result('deleteSlide',result,params={'slideId':256,'expectedToken':'before'})
        with self.assertRaises(ContractValidationError): contracts.validate_result('listSlides',{'slides':[slide,slide],'token':'x'},params={})
