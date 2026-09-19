import unittest


class DesignValuePlanTests(unittest.TestCase):
    def test_native_plans_are_valid_and_isolate_documents_by_full_root(self):
        from diagnostics.design_value.plans import seed, edit
        from wps_skills.client.applications import compile_request, contracts_for
        from wps_skills.client.task_client import _preflight
        for app in ('word', 'excel', 'ppt'):
            first = seed(app, 'C:/experiments/one/app', 'sample')
            second = seed(app, 'C:/experiments/two/app', 'sample')
            self.assertNotEqual(first.path, second.path)
            for case in (first, second, edit(app, 'C:/experiments/one/app', 'edit', first.path)):
                _preflight(compile_request(case.request, app), contracts_for(app))
            self.assertTrue(edit(app, 'C:/experiments/one/app', 'edit', first.path).request['completion'])


if __name__ == '__main__':
    unittest.main()
