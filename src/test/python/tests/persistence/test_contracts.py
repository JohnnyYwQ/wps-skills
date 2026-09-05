import copy
import unittest
from wps_skills.core.action_session import ContractValidationError
from wps_skills.word.contracts import WORD_TARGET_CONTRACT_SET
from wps_skills.excel.contracts import EXCEL_TARGET_CONTRACT_SET
from wps_skills.ppt.contracts import PPT_TARGET_CONTRACT_SET


class PersistenceContractTests(unittest.TestCase):
    def test_create_has_no_path_or_implicit_persistence(self):
        for contracts,action in ((WORD_TARGET_CONTRACT_SET,'createDocument'),(EXCEL_TARGET_CONTRACT_SET,'createWorkbook'),(PPT_TARGET_CONTRACT_SET,'createPresentation')):
            with self.subTest(app=contracts.application):
                contracts.validate_params(action,{})
                with self.assertRaises(ContractValidationError): contracts.validate_params(action,{'path':'C:/work/new.xlsx'})

    def test_save_as_requires_absent_destination_policy_and_exact_format(self):
        for contracts,ext in ((WORD_TARGET_CONTRACT_SET,'docx'),(EXCEL_TARGET_CONTRACT_SET,'xlsx'),(PPT_TARGET_CONTRACT_SET,'pptx')):
            params={'outputPath':'C:/work/new.'+ext,'overwritePolicy':'failIfExists'}
            contracts.validate_params('saveAs',params)
            for patch in ({'overwritePolicy':'replace'},{'outputPath':'new.'+ext},{'outputPath':'C:/work/new.pdf'},{'outputPath':'C:/work/../new.'+ext}):
                with self.subTest(app=contracts.application,patch=patch),self.assertRaises(ContractValidationError):
                    contracts.validate_params('saveAs',dict(params,**patch))
            with self.assertRaises(ContractValidationError):contracts.validate_params('saveAs',{'outputPath':params['outputPath']})

    def test_export_cannot_claim_saved_state_for_an_unsaved_document(self):
        for contracts in (EXCEL_TARGET_CONTRACT_SET,PPT_TARGET_CONTRACT_SET):
            params={'outputPath':'C:/work/new.pdf','overwritePolicy':'failIfExists'}
            state={'persistenceState':'unsaved','readOnly':False}
            data={'artifact':{'path':params['outputPath'],'format':'pdf','sizeBytes':100},'documentStateBefore':state,'documentStateAfter':dict(state),'replacedExisting':False}
            contracts.validate_result('exportPdf',data,params=params)
            altered=copy.deepcopy(data);altered['documentStateAfter']['persistenceState']='saved'
            with self.assertRaises(ContractValidationError): contracts.validate_result('exportPdf',altered,params=params)
            altered=copy.deepcopy(data);altered['artifact']['path']='C:/work/other.pdf'
            with self.assertRaises(ContractValidationError): contracts.validate_result('exportPdf',altered,params=params)

    def test_png_requires_observed_slide_and_bounded_pixel_dimensions(self):
        params={'outputPath':'C:/work/slide.png','overwritePolicy':'failIfExists','slideId':256,'width':1280,'height':720}
        PPT_TARGET_CONTRACT_SET.validate_params('exportSlideImage',params)
        for patch in ({'width':0},{'height':4097},{'width':True},{'slideId':0},{'outputPath':'C:/work/slide.jpg'}):
            with self.assertRaises(ContractValidationError):PPT_TARGET_CONTRACT_SET.validate_params('exportSlideImage',dict(params,**patch))
