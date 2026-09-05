import copy
import unittest

from wps_skills.core.action_session import ContractValidationError
from wps_skills.excel.contracts import EXCEL_TARGET_CONTRACT_SET as contracts, range_bounds


class ExcelContractsTests(unittest.TestCase):
    def test_examples_and_index_share_authority(self):
        for contract in contracts.contracts:
            for example in contract.examples:
                contracts.validate_params(contract.name, example['params'])
        self.assertEqual(34, len(contracts.action_index()))
        self.assertEqual('establish', contracts.resolve('openWorkbook').binding_role)

    def test_ranges_are_bounded_sheet_local_rectangles(self):
        self.assertEqual((1, 1, 20, 3), range_bounds('A1:C20'))
        self.assertEqual((1048576, 16384, 1048576, 16384), range_bounds('XFD1048576'))
        for address in ('A:A', 'A1;B2', 'Sheet1!A1', '[other.xlsx]A1', 'a1', '$A$1', 'A0', 'B2:A1', 'A1:A1001', 'XFE1', 'A1048577'):
            with self.subTest(address=address), self.assertRaises(ValueError):
                range_bounds(address)

    def test_writes_reject_bad_shapes_and_missing_tokens(self):
        base = {'sheet': 'Sheet1', 'address': 'A1:B2', 'expectedToken': 'observed', 'values': [[1, 2], [3, 4]]}
        contracts.validate_params('writeRange', base)
        for patch in ({'values': [[1, 2]]}, {'values': [[1], [2, 3]]}, {'values': [[1, 2], [3, float('nan')]]}, {'expectedToken': ''}, {'activeWorkbook': True}):
            with self.subTest(patch=patch), self.assertRaises(ContractValidationError):
                contracts.validate_params('writeRange', dict(base, **patch))
        del base['expectedToken']
        with self.assertRaises(ContractValidationError):
            contracts.validate_params('writeRange', base)

    def test_only_explicit_windows_xlsx_locators_are_accepted(self):
        for path in (r'C:\work\book.xlsx', r'\\server\share\book.xlsx'):
            contracts.validate_params('openWorkbook', {'path': path})
        for path in ('book.xlsx', '/tmp/book.xlsx', r'C:\book.xlsm', r'C:\NUL.xlsx', r'C:\bad.\book.xlsx', 'C:\\bad\n.xlsx'):
            with self.subTest(path=path), self.assertRaises(ContractValidationError):
                contracts.validate_params('openWorkbook', {'path': path})

    def test_formulas_exclude_external_calls_names_and_other_sheets(self):
        base = {'sheet': 'Sheet1', 'address': 'A1', 'expectedToken': 'observed'}
        for formula in ('=SUM(B1:B4)', '=IF(B1>0,"中文","")', '=ROUND($B$2*1.1,2)', '=IFERROR(1/0,0)'):
            contracts.validate_params('setFormulas', dict(base, formulas=[[formula]]))
        for formula in ('=WEBSERVICE("https://example.com")', '=[book.xlsx]Sheet1!A1', '=Sheet2!A1', '=cmd|abc!A1', '=UnknownName', '=SUM(XFE1)', '=SEQUENCE(3)'):
            with self.subTest(formula=formula), self.assertRaises(ContractValidationError):
                contracts.validate_params('setFormulas', dict(base, formulas=[[formula]]))

    def test_result_validation_checks_actual_values_types_and_scope(self):
        params = {'sheet': 'Sheet1', 'address': 'A1', 'expectedToken': 'before', 'values': [[True]]}
        good = {'sheet': 'Sheet1', 'address': 'A1', 'token': 'after', 'cells': [[{
            'value': True, 'formula': None, 'text': 'TRUE', 'errorCode': None, 'numberFormat': 'General', 'bold': False}]]}
        contracts.validate_result('writeRange', good, params=params)
        for patch in ({'value': 1}, {'formula': '=1'}, {'errorCode': 2007}):
            bad = copy.deepcopy(good)
            bad['cells'][0][0].update(patch)
            with self.subTest(patch=patch), self.assertRaises(ContractValidationError):
                contracts.validate_result('writeRange', bad, params=params)
        bad = dict(good, sheet='Other')
        with self.assertRaises(ContractValidationError):
            contracts.validate_result('writeRange', bad, params=params)

    def test_partial_format_is_nonempty_and_result_must_match(self):
        with self.assertRaises(ContractValidationError):
            contracts.validate_params('formatRange', {'sheet': 'Sheet1', 'address': 'A1', 'expectedToken': 'x', 'format': {}})

    def test_worksheet_pages_cannot_claim_unobserved_rows(self):
        params = {'offset': 0, 'limit': 1}
        good = {'total': 2, 'nextOffset': 1, 'worksheets': [{'name': 'Sheet1', 'index': 1}]}
        contracts.validate_result('listWorksheets', good, params=params)
        with self.assertRaises(ContractValidationError):
            contracts.validate_result('listWorksheets', dict(good, nextOffset=None), params=params)

    def test_common_actions_reject_unsafe_scopes_and_invalid_names(self):
        base = {'sheet': 'Sheet1', 'expectedToken': 'observed'}
        for name in ('bad/name', "'quoted", 'control\x00', 'a'*32):
            with self.subTest(name=name), self.assertRaises(ContractValidationError):
                contracts.validate_params('renameWorksheet', dict(base, name=name))
        for action, patch in (
            ('insertRows', {'start': 1048576, 'count': 2}),
            ('deleteColumns', {'start': 16384, 'count': 2}),
            ('sortRange', {'address': 'A1:B3', 'column': 3, 'order': 'ascending', 'header': True}),
            ('filterRange', {'address': 'A1:B1', 'column': 1, 'value': 'x'}),
            ('copyRange', {'address': 'A1:B2', 'targetSheet': 'Sheet1', 'targetAddress': 'B2:C3', 'targetToken': 't'}),
            ('copyRange', {'address': 'A1:B2', 'targetSheet': 'Other', 'targetAddress': 'A1', 'targetToken': 't'}),
        ):
            with self.subTest(action=action, patch=patch), self.assertRaises(ContractValidationError):
                contracts.validate_params(action, dict(base, **patch))

    def test_find_result_cannot_escape_region_or_fabricate_matches(self):
        params = {'sheet': 'Sheet1', 'address': 'A1:B2', 'text': 'one', 'matchCase': False, 'wholeCell': True, 'lookIn': 'values'}
        result = {'sheet': 'Sheet1', 'address': 'A1:B2', 'token': 't', 'matches': [{'address': 'A1', 'text': 'ONE'}]}
        contracts.validate_result('findInRange', result, params=params)
        for matches in ([{'address': 'C1', 'text': 'one'}], [{'address': 'A1', 'text': 'two'}], result['matches'] * 2):
            with self.subTest(matches=matches), self.assertRaises(ContractValidationError):
                contracts.validate_result('findInRange', dict(result, matches=matches), params=params)

    def test_structure_result_proves_requested_name_and_position(self):
        params = {'sheet': 'Sheet1', 'expectedToken': 'before', 'name': 'Renamed'}
        result = {'name': 'Renamed', 'index': 1, 'visible': True, 'usedAddress': 'A1', 'token': 'after'}
        contracts.validate_result('renameWorksheet', result, params=params)
        with self.assertRaises(ContractValidationError):
            contracts.validate_result('renameWorksheet', dict(result, name='Other'), params=params)
        with self.assertRaises(ContractValidationError):
            contracts.validate_result('moveWorksheet', result, params={'sheet':'Renamed', 'expectedToken':'t', 'index':2})
