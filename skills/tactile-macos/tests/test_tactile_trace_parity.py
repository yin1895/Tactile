import importlib.util
import sys
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[3]
MAC_TRACE_PATH = REPO_ROOT / "skills" / "tactile-macos" / "scripts" / "utils" / "tactile_trace.py"
WINDOWS_TRACE_PATH = REPO_ROOT / "skills" / "tactile-windows" / "scripts" / "utils" / "tactile_trace.py"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


mac_trace = load_module("tactile_trace_parity_macos", MAC_TRACE_PATH)
windows_trace = load_module("tactile_trace_parity_windows", WINDOWS_TRACE_PATH)


class TactileTraceParityTests(unittest.TestCase):
    def test_helpers_expose_the_same_trace_surface(self):
        expected = {
            "build_trace",
            "build_fast_path_trace",
            "trace_summary",
            "replay_trace_payloads",
            "replay_trace_files",
            "source_name",
        }

        for name in expected:
            self.assertTrue(callable(getattr(mac_trace, name, None)), name)
            self.assertTrue(callable(getattr(windows_trace, name, None)), name)

    def test_core_outputs_keep_matching_keys(self):
        run_log = {
            "target": {"identifier": "Calculator"},
            "instruction": "click via OCR",
            "task_source": "adapter_dry_run",
            "final_status": "finished",
            "steps": [
                {
                    "step": 1,
                    "plan": {
                        "summary": "planned route",
                        "actions": [{"type": "route", "source": "ocr_coordinate"}],
                    },
                    "execution_results": [
                        {
                            "ok": True,
                            "mode": "ocr_coordinate",
                            "action": {"type": "route", "source": "ocr_coordinate"},
                        }
                    ],
                    "verification": {"status": "planned", "covered": True},
                }
            ],
        }

        mac = mac_trace.build_trace(run_log, platform="macos")
        windows = windows_trace.build_trace(run_log, platform="windows")
        mac_replay = mac_trace.replay_trace_payloads([mac])
        windows_replay = windows_trace.replay_trace_payloads([windows])

        self.assertEqual(set(mac.keys()), set(windows.keys()))
        self.assertEqual(set(mac["metrics"].keys()), set(windows["metrics"].keys()))
        self.assertEqual(set(mac["outcome"].keys()), set(windows["outcome"].keys()))
        self.assertEqual(set(mac_replay.keys()), set(windows_replay.keys()))
        self.assertEqual(mac_replay["planned_coordinate_sources"], windows_replay["planned_coordinate_sources"])


if __name__ == "__main__":
    unittest.main()
