import unittest

from wps_skills.windows.bridge_types import BackendActionFailure


class BackendActionFailureTests(unittest.TestCase):
    def test_failure_shape_is_closed(self):
        failure = BackendActionFailure(
            outcome="unknown",
            code="OUTPUT_WRITE_FAILED",
            message="write result is unknown",
            binding_disposition="unprovable",
        )
        self.assertEqual("unknown", failure.outcome)
        self.assertEqual("unprovable", failure.binding_disposition)

        with self.assertRaisesRegex(ValueError, "cannot prove binding loss"):
            BackendActionFailure(
                outcome="unknown",
                code="OUTPUT_WRITE_FAILED",
                message="unknown",
                binding_disposition="lost",
            )


