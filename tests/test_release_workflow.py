from pathlib import Path
import unittest


WORKFLOW_PATH = (
    Path(__file__).resolve().parents[1]
    / ".github"
    / "workflows"
    / "release-exe.yml"
)


class ReleaseWorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.workflow = WORKFLOW_PATH.read_text(encoding="utf-8")

    def test_release_tag_is_resolved_without_shell_expression_injection(self):
        resolve_start = self.workflow.index("      - name: Resolve release tag")
        checkout_start = self.workflow.index("      - name: Checkout release tag")
        resolve_step = self.workflow[resolve_start:checkout_start]
        run_script = resolve_step[resolve_step.index("        run: |") :]

        self.assertIn("RELEASE_TAG_INPUT: ${{", resolve_step)
        self.assertIn("$env:RELEASE_TAG_INPUT", run_script)
        self.assertNotIn("${{", run_script)

    def test_checkout_uses_the_validated_release_tag(self):
        resolve_start = self.workflow.index("      - name: Resolve release tag")
        checkout_start = self.workflow.index("      - name: Checkout release tag")
        setup_start = self.workflow.index("      - name: Set up Python")
        checkout_step = self.workflow[checkout_start:setup_start]

        self.assertLess(resolve_start, checkout_start)
        self.assertIn("ref: ${{ steps.release.outputs.tag }}", checkout_step)
        self.assertIn("persist-credentials: false", checkout_step)

    def test_missing_or_empty_test_directory_blocks_release(self):
        tests_start = self.workflow.index("      - name: Run tests")
        build_start = self.workflow.index("      - name: Build EXE with PyInstaller")
        tests_step = self.workflow[tests_start:build_start]

        self.assertIn('Test-Path -LiteralPath "tests" -PathType Container', tests_step)
        self.assertIn('Get-ChildItem -LiteralPath "tests"', tests_step)
        self.assertGreaterEqual(tests_step.count("throw "), 2)
        self.assertNotIn("skipping tests", tests_step)

    def test_release_bundles_certifi_ca_data_for_wss(self):
        verify_start = self.workflow.index("      - name: Verify build inputs")
        tests_start = self.workflow.index("      - name: Run tests")
        verify_step = self.workflow[verify_start:tests_start]
        build_start = self.workflow.index("      - name: Build EXE with PyInstaller")
        package_start = self.workflow.index("      - name: Package release files")
        build_step = self.workflow[build_start:package_start]

        self.assertIn("import certifi", verify_step)
        self.assertIn('--collect-data "certifi"', build_step)


if __name__ == "__main__":
    unittest.main()
