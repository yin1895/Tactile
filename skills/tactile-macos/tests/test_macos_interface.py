import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "scripts" / "macos_interface.py"
SPEC = importlib.util.spec_from_file_location("macos_interface", SCRIPT_PATH)
assert SPEC is not None
assert SPEC.loader is not None
macos_interface = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(macos_interface)


class MacosInterfaceOcrTests(unittest.TestCase):
    def test_launch_debug_ax_grid_uses_highlight_tool_without_stdout_pollution(self):
        calls = []

        class FakePopen:
            pass

        def fake_ensure_products(repo, products):
            calls.append(("ensure", repo, products))

        def fake_popen(cmd, *, cwd, text, stdout, stderr):
            calls.append(("popen", cmd, cwd, text, stdout, stderr))
            return FakePopen()

        original_ensure_products = macos_interface.ensure_products
        original_popen = macos_interface.subprocess.Popen
        try:
            macos_interface.ensure_products = fake_ensure_products
            macos_interface.subprocess.Popen = fake_popen

            proc = macos_interface.launch_debug_ax_grid(Path("/tmp/repo"), 12345, 0.5)
        finally:
            macos_interface.ensure_products = original_ensure_products
            macos_interface.subprocess.Popen = original_popen

        self.assertIsInstance(proc, FakePopen)
        self.assertEqual(calls[0], ("ensure", Path("/tmp/repo"), ["HighlightTraversalTool"]))
        popen_call = calls[1]
        self.assertEqual(popen_call[0], "popen")
        self.assertEqual(popen_call[1], ["/tmp/repo/.build/debug/HighlightTraversalTool", "12345", "--no-activate", "--duration", "0.5"])
        self.assertEqual(popen_call[2], Path("/tmp/repo"))
        self.assertIs(popen_call[4], subprocess.DEVNULL)
        self.assertIs(popen_call[5], subprocess.DEVNULL)

    def test_add_screen_frames_maps_retina_image_pixels_to_screen_points(self):
        payload = {
            "imageWidth": 200,
            "imageHeight": 100,
            "lines": [
                {
                    "text": "Search",
                    "confidence": 1,
                    "frame": {"x": 20, "y": 10, "width": 40, "height": 20},
                }
            ],
        }

        macos_interface.add_screen_frames_to_ocr_payload(payload, (100, 200, 100, 50))

        line = payload["lines"][0]
        self.assertEqual(
            payload["coordinateSpace"],
            {
                "frame": "image_pixels_relative_to_screenshot",
                "screenFrame": "screen_points_top_left",
            },
        )
        self.assertEqual(line["frame"], {"x": 20, "y": 10, "width": 40, "height": 20})
        self.assertEqual(line["imageFrame"], line["frame"])
        self.assertEqual(line["screenFrame"], {"x": 110, "y": 205, "width": 20, "height": 10})
        self.assertEqual(line["screenCenter"], {"x": 120, "y": 210})

    def test_ocr_tsv_prefers_clickable_screen_frame_when_present(self):
        payload = {
            "lines": [
                {
                    "text": "Search",
                    "confidence": 0.9,
                    "frame": {"x": 20, "y": 10, "width": 40, "height": 20},
                    "screenFrame": {"x": 110, "y": 205, "width": 20, "height": 10},
                }
            ],
        }

        self.assertEqual(macos_interface.format_ocr_payload(payload, "tsv"), "110\t205\t20\t10\t0.90\tSearch\n")

    def test_image_only_ocr_keeps_existing_image_frame_tsv(self):
        payload = {
            "lines": [
                {
                    "text": "Search",
                    "confidence": 0.9,
                    "frame": {"x": 20, "y": 10, "width": 40, "height": 20},
                }
            ],
        }

        macos_interface.add_screen_frames_to_ocr_payload(payload, None)

        self.assertNotIn("coordinateSpace", payload)
        self.assertNotIn("screenFrame", payload["lines"][0])
        self.assertEqual(macos_interface.format_ocr_payload(payload, "tsv"), "20\t10\t40\t20\t0.90\tSearch\n")

    def test_session_artifact_dir_uses_explicit_session_directory(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            session_dir = Path(temp_dir) / "session-1"

            artifact_dir = macos_interface.session_artifact_dir(
                env={"TACTILE_SESSION_DIR": str(session_dir)},
            )

            self.assertEqual(artifact_dir, session_dir / "macos-app-workflow")
            self.assertTrue(artifact_dir.is_dir())

    def test_session_artifact_dir_uses_matching_tactile_session_scope(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            workspace = Path(temp_dir)
            session_scope = workspace / ".claw" / "sessions" / "scope-a"
            session_scope.mkdir(parents=True)
            (session_scope / "session-123.jsonl").write_text("{}", encoding="utf-8")

            artifact_dir = macos_interface.session_artifact_dir(
                cwd=workspace,
                env={"TACTILE_SESSION_ID": "session-123"},
            )

            self.assertEqual(artifact_dir, session_scope.resolve() / "macos-app-workflow")
            self.assertTrue(artifact_dir.is_dir())

    def test_temp_output_path_is_relocated_to_session_artifacts(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            session_dir = Path(temp_dir) / "session-1"
            original_env = dict(macos_interface.os.environ)
            try:
                macos_interface.os.environ.clear()
                macos_interface.os.environ.update({"TACTILE_SESSION_DIR": str(session_dir)})

                output = macos_interface.session_scoped_output_path(Path("/tmp/lark_switch_org_run.json"))
            finally:
                macos_interface.os.environ.clear()
                macos_interface.os.environ.update(original_env)

            self.assertEqual(output, session_dir / "macos-app-workflow" / "lark_switch_org_run.json")

    def test_direct_output_path_preserves_temporary_path(self):
        output = macos_interface.resolved_output_path(Path("/tmp/lark_switch_org_run.json"), direct_output=True)

        self.assertEqual(output, Path("/tmp/lark_switch_org_run.json"))

    def test_var_folders_temp_output_path_is_relocated_to_session_artifacts(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            session_dir = Path(temp_dir) / "session-1"
            original_env = dict(macos_interface.os.environ)
            original_gettempdir = macos_interface.tempfile.gettempdir
            try:
                macos_interface.os.environ.clear()
                macos_interface.os.environ.update({"TACTILE_SESSION_DIR": str(session_dir)})
                macos_interface.tempfile.gettempdir = lambda: "/var/folders/sl/test/T"

                output = macos_interface.session_scoped_output_path(Path("/var/folders/sl/test/T/opencode/feishu_workflow.json"))
            finally:
                macos_interface.tempfile.gettempdir = original_gettempdir
                macos_interface.os.environ.clear()
                macos_interface.os.environ.update(original_env)

            self.assertEqual(output, session_dir / "macos-app-workflow" / "feishu_workflow.json")

    def test_legacy_dot_artifact_output_path_is_relocated_to_session_artifacts(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            session_dir = Path(temp_dir) / "session-1"
            legacy_output = Path(temp_dir) / "Desktop" / ".macos-app-workflow" / "feishu_directory_send_log.json"
            original_env = dict(macos_interface.os.environ)
            try:
                macos_interface.os.environ.clear()
                macos_interface.os.environ.update({"TACTILE_SESSION_DIR": str(session_dir)})

                output = macos_interface.session_scoped_output_path(legacy_output)
            finally:
                macos_interface.os.environ.clear()
                macos_interface.os.environ.update(original_env)

            self.assertEqual(output, session_dir / "macos-app-workflow" / "feishu_directory_send_log.json")

    def test_session_artifact_dir_uses_opencode_session_artifacts_fallback(self):
        original_env = dict(macos_interface.os.environ)
        try:
            macos_interface.os.environ.clear()
            macos_interface.os.environ.update({"TACTILE_SESSION_ID": "session-abc"})

            artifact_dir = macos_interface.session_artifact_dir(cwd=Path("/tmp"), create=False)
        finally:
            macos_interface.os.environ.clear()
            macos_interface.os.environ.update(original_env)

        self.assertEqual(
            artifact_dir,
            Path.home() / ".local" / "share" / "opencode" / "storage" / "session_artifacts" / "session-abc" / "macos-app-workflow",
        )

    def test_python_helper_creates_uv_managed_venv(self):
        workflow_dir = Path(__file__).resolve().parents[1]
        helper = workflow_dir / "bin" / "tactile-macos-python"
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            fake_bin = temp_path / "bin"
            fake_bin.mkdir()
            fake_uv = fake_bin / "uv"
            fake_uv_log = temp_path / "uv.log"
            fake_uv.write_text(
                "#!/bin/sh\n"
                "set -eu\n"
                "for arg in \"$@\"; do printf '<%s>' \"$arg\" >>\"$FAKE_UV_LOG\"; done\n"
                "printf '\\n' >>\"$FAKE_UV_LOG\"\n"
                "venv=${UV_PROJECT_ENVIRONMENT:?}\n"
                "mkdir -p \"$venv/bin\"\n"
                "cat >\"$venv/bin/python\" <<'PY'\n"
                "#!/bin/sh\n"
                "printf 'fake-python'\n"
                "for arg in \"$@\"; do printf ' <%s>' \"$arg\"; done\n"
                "printf '\\n'\n"
                "PY\n"
                "chmod +x \"$venv/bin/python\"\n",
                encoding="utf-8",
            )
            fake_uv.chmod(0o755)
            venv = temp_path / "workflow-venv"
            env = dict(os.environ)
            env.update(
                {
                    "PATH": os.pathsep.join([os.fspath(fake_bin), env.get("PATH", "")]),
                    "FAKE_UV_LOG": os.fspath(fake_uv_log),
                    "TACTILE_MACOS_WORKFLOW_VENV": os.fspath(venv),
                    "TACTILE_MACOS_PYTHON_VERSION": "3.12",
                }
            )

            proc = subprocess.run(
                [os.fspath(helper), "-c", "print('ignored')"],
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                env=env,
                check=True,
            )

            self.assertTrue((venv / "bin" / "python").is_file())
            self.assertTrue((venv / ".tactile-deps-stamp").is_file())
            self.assertIn(
                f"<sync><--project><{workflow_dir}><--no-dev><--frozen><--python><3.12>",
                fake_uv_log.read_text(encoding="utf-8"),
            )
            self.assertEqual(proc.stdout, "fake-python <-c> <print('ignored')>\n")

    def test_wrapper_uses_system_python_for_stdlib_commands_without_uv(self):
        workflow_dir = Path(__file__).resolve().parents[1]
        wrapper = workflow_dir / "bin" / "tactile-macos"
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "doctor.json"
            env = dict(os.environ)
            env.update(
                {
                    "PATH": "/usr/bin:/bin",
                    "TACTILE_MACOS_SYSTEM_PYTHON": sys.executable,
                    "TACTILE_UV": "",
                }
            )

            proc = subprocess.run(
                [os.fspath(wrapper), "doctor", "--output", os.fspath(output), "--direct-output"],
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                env=env,
                check=True,
            )

            self.assertEqual(proc.stdout, f"{output}\n")
            self.assertTrue(output.is_file())
            payload = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(payload["python"], sys.executable)
