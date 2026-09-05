import copy
import unittest
from wps_skills.core.action_session import ContractValidationError,ControllerResult
from wps_skills.ppt.contracts import PPT_TARGET_CONTRACT_SET as contracts, PPT_PRODUCTION_CONTRACT_SET


def shape(id=2,left=10,width=40,top=20,height=30):
    return {'id':id,'name':'A','type':1,'left':left,'top':top,'width':width,'height':height,'rotation':0,
            'zOrder':id-1,'autoShapeType':1,'text':'Hello','font':None}


def slide(shapes):
    return {'slide':{'id':256,'index':1,'name':'Slide1','layout':12,'shapeCount':len(shapes)},'shapes':shapes,'token':'after'}


class PptCommonContractsTests(unittest.TestCase):
    def test_admitted_common_actions_have_valid_examples_and_one_authority(self):
        self.assertEqual(33,len(contracts.contracts))
        for c in contracts.contracts:
            for example in c.examples:contracts.validate_params(c.name,example['params'])
        self.assertEqual(33,len(PPT_PRODUCTION_CONTRACT_SET.contracts))

    def test_base_snapshots_accept_extended_native_shape_observations(self):
        contracts.validate_result('getSlideInfo',slide([shape()]),params={'slideId':256})
        PPT_PRODUCTION_CONTRACT_SET.validate_result('getSlideInfo',slide([shape()]),params={'slideId':256})

    def test_collection_and_settings_parameters_cannot_be_ambiguous(self):
        for action,params in (
            ('alignShapes',{'shapeIds':[2,2],'alignment':'left'}),
            ('distributeShapes',{'shapeIds':[2,3],'direction':'horizontal'}),
            ('setSlideSettings',{'settings':{'followMasterBackground':True,'backgroundColor':255}}),
            ('renameShape',{'shapeId':2,'name':'bad\nname'}),
            ('formatShape',{'shapeId':2,'format':{'lineColor':255,'lineVisible':False}}),
            ('formatShape',{'shapeId':2,'format':{'fillTransparency':0.5,'fillVisible':False}}),
            ('replaceText',{'shapeId':2,'find':'','replacement':'new'}),
        ):
            with self.subTest(action=action),self.assertRaises(ContractValidationError):
                contracts.validate_params(action,dict(slideId=256,expectedToken='before',**params))

    def test_image_locators_and_table_matrices_are_bounded(self):
        base={'slideId':256,'expectedToken':'before','left':0,'top':0,'width':200,'height':100}
        contracts.validate_params('addImage',dict(base,path=r'C:\images\photo.JPG'))
        for path in ('relative.png',r'C:\NUL.png',r'C:\photo.svg','https://host/image.png'):
            with self.assertRaises(ContractValidationError):contracts.validate_params('addImage',dict(base,path=path))
        for values in ([],[['A'],['B','C']],[['x']*10]*11,[['x'*2001]],[['😀'*2000]*6]):
            with self.subTest(rows=len(values)),self.assertRaises(ContractValidationError):contracts.validate_params('addTable',dict(base,values=values))

    def test_frozen_table_result_proves_exact_values_and_shape(self):
        params={'slideId':256,'shapeId':5,'expectedToken':'before','values':[['A','😀']]}
        data={'slideId':256,'shapeId':5,'rows':1,'columns':2,'values':[['A','😀']],'token':'after'}
        frozen=ControllerResult.succeeded(data=data,controller_state='usable',binding_disposition='unchanged').data
        contracts.validate_result('writeTable',frozen,params=params)
        for patch in ({'shapeId':6},{'rows':2},{'values':[['A','wrong']]}):
            with self.assertRaises(ContractValidationError):contracts.validate_result('writeTable',dict(data,**patch),params=params)

    def test_alignment_distribution_and_stacking_must_be_observed(self):
        params={'slideId':256,'expectedToken':'before','shapeIds':[2,3],'alignment':'left'}
        good=slide([shape(2,10),shape(3,10)])
        contracts.validate_result('alignShapes',good,params=params)
        with self.assertRaises(ContractValidationError):contracts.validate_result('alignShapes',slide([shape(2,10),shape(3,20)]),params=params)
        params=dict(slideId=256,expectedToken='before',shapeIds=[2,3,4],direction='horizontal')
        good=slide([shape(2,0,40),shape(3,60,50),shape(4,130,70)])
        contracts.validate_result('distributeShapes',good,params=params)
        bad=copy.deepcopy(good);bad['shapes'][1]['left']=70
        with self.assertRaises(ContractValidationError):contracts.validate_result('distributeShapes',bad,params=params)
        with self.assertRaises(ContractValidationError):contracts.validate_result('setShapeOrder',good,params={'slideId':256,'shapeId':2,'expectedToken':'before','position':'front'})

    def test_notes_and_find_readback_cannot_fabricate_content(self):
        notes={'slideId':256,'text':'wrong','token':'after'}
        with self.assertRaises(ContractValidationError):contracts.validate_result('setSlideNotes',notes,params={'slideId':256,'text':'wanted','expectedToken':'before'})
        params={'slideId':256,'text':'😀'}
        match={'shapeId':2,'start':0,'length':2,'text':'😀'}
        result={'slideId':256,'matches':[match],'token':'observed'}
        contracts.validate_result('findText',result,params=params)
        with self.assertRaises(ContractValidationError):contracts.validate_result('findText',dict(result,matches=[dict(match,length=1)]),params=params)
        with self.assertRaises(ContractValidationError):contracts.validate_result('findText',dict(result,matches=[match,match]),params=params)

    def test_paragraph_spacing_must_be_points_on_every_paragraph(self):
        paragraph={'alignment':2,'spaceBefore':6,'spaceAfter':9,'spaceBeforeInLines':False,'spaceAfterInLines':False,'bulletVisible':True,'bulletType':1}
        data={'slideId':256,'shape':shape(),'appearance':{'fillVisible':True,'fillType':1,'fillColor':255,'fillTransparency':0.2,'lineVisible':False,'lineColor':0,'lineWidth':1},
              'textBox':None,'paragraphs':[paragraph,paragraph],'token':'after'}
        params={'slideId':256,'shapeId':2,'expectedToken':'before','format':{'alignment':'center','spaceBefore':6}}
        contracts.validate_result('formatParagraph',data,params=params)
        for patch in ({'spaceBeforeInLines':True},{'alignment':1}):
            bad=copy.deepcopy(data);bad['paragraphs'][1].update(patch)
            with self.assertRaises(ContractValidationError):contracts.validate_result('formatParagraph',bad,params=params)
