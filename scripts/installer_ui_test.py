"""Installer-only artwork, animation bounds and preview safeguards."""
from __future__ import annotations

import unittest
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]


class InstallerStoryTests(unittest.TestCase):
    def test_sprite_sheet_is_transparent_and_bounded(self):
        path = ROOT / "installer/mascot-poses.png"
        self.assertLess(path.stat().st_size, 2_000_000)
        with Image.open(path) as sprite:
            self.assertEqual(sprite.mode, "RGBA")
            self.assertEqual(sprite.width, sprite.height)
            self.assertEqual(sprite.width % 2, 0)
            self.assertEqual(sprite.getextrema()[-1], (0, 255))
            self.assertGreaterEqual(sprite.width, 1024)

    def test_animation_has_motion_preference_pause_and_cleanup(self):
        story = (ROOT / "installer/story.iss").read_text(encoding="utf-8")
        for marker in ("SPI_GETCLIENTAREAANIMATION = $1042", "AllowMotion := False",
                       "StoryTick >= 100", "SetTimer(0, 0, 80,", "GetForegroundWindow() <> WizardForm.Handle",
                       "KillTimer(0, StoryTimer)", "procedure DeinitializeSetup()", "not StoryMotion.Checked"):
            self.assertIn(marker, story)
        self.assertIn("hWnd: HWND; nIDEvent: UINT_PTR", story)
        self.assertIn("CurProgress * 100.0", story)
        self.assertIn("if WizardSilent() then Exit", story)
        self.assertIn("WizardForm.AdjustLabelHeight(WizardForm.FinishedLabel)", story)
        self.assertIn("WizardForm.RunList.Top := WizardForm.FinishedLabel.Top + WizardForm.FinishedLabel.Height", story)
        timer_body = story.split("procedure StoryTimerTick", 1)[1].split("procedure StartStoryMotion", 1)[0]
        self.assertNotIn("ProgressGauge", timer_body)
        self.assertNotIn("LoadFromFile", timer_body)

    def test_artwork_is_installer_only_and_preview_cannot_install(self):
        installer = (ROOT / "installer/zadoo.iss").read_text(encoding="utf-8")
        self.assertIn('Source: "mascot-poses.png"; Flags: dontcopy', installer)
        self.assertIn('#include "story.iss"', installer)
        self.assertIn("PrivilegesRequired=admin", installer)
        self.assertNotIn("#define ZadooVisualPreview", installer)
        preview = (ROOT / "scripts/preview_installer.ps1").read_text(encoding="utf-8")
        self.assertIn("Result := CurPageID <> wpReady", preview)
        self.assertIn("Uninstallable=no", preview)
        self.assertIn("CreateAppDir=no", preview)
        self.assertIn("Unsafe preview source sections", preview)
        self.assertFalse((ROOT / "zadoo_vnc/static/ui/mascot-poses.png").exists())

    def test_silent_uninstall_preserves_settings_without_prompting(self):
        installer = (ROOT / "installer/zadoo.iss").read_text(encoding="utf-8")
        uninstall = installer.split("function InitializeUninstall(): Boolean;", 1)[1]
        silent_branch = uninstall.split("if UninstallSilent() then", 1)[1].split("end;", 1)[0]
        self.assertIn("Exit;", silent_branch)
        self.assertNotIn("DelTree", silent_branch)
        self.assertLess(uninstall.index("if UninstallSilent() then"), uninstall.index("DirExists("))
        self.assertIn("SuppressibleMsgBox(", uninstall)
        self.assertIn("MB_YESNO or MB_DEFBUTTON2, IDNO) = IDYES", uninstall)

    def test_real_lifecycle_harness_is_disposable_runner_only(self):
        harness = (ROOT / "scripts/installer_e2e_test.ps1").read_text(encoding="utf-8")
        for marker in ("$env:GITHUB_ACTIONS -ne 'true'", "$env:RUNNER_ENVIRONMENT -ne 'github-hosted'",
                       "$env:RUNNER_OS -ne 'Windows'", "Assert-TestInstallPath", "ReparsePoint",
                       "requires a pristine runner", "refuses to replace an existing Zadoo task",
                       "silent-uninstall-without-suppression", "very-silent-uninstall-with-suppression",
                       "Assert-DataPreserved"):
            self.assertIn(marker, harness)
        workflow = (ROOT / ".github/workflows/windows-x64-ci.yml").read_text(encoding="utf-8")
        self.assertIn("44567571c7a122d76ddf21d11265ddf2269908d6", workflow)
        self.assertIn("scripts\\installer_e2e_test.ps1 -PreviousInstaller", workflow)


if __name__ == "__main__":
    unittest.main()
