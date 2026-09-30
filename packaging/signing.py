"""Opt-in release signing using credentials already provisioned on the build host."""

import os
from pathlib import Path
import shutil
import subprocess


def setting(name):
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Signed release requires {name}; refusing unsigned fallback.")
    return value


def check_configuration(target):
    if target == "macos":
        identity = setting("APPLE_SIGNING_IDENTITY")
        if not identity.startswith("Developer ID Application:"):
            raise RuntimeError("Use a Developer ID Application identity, not an ad-hoc identity.")
        setting("APPLE_NOTARY_PROFILE")
    elif target == "windows":
        setting("WINDOWS_CERT_SHA1")
        setting("WINDOWS_TIMESTAMP_URL")
        if not shutil.which(os.environ.get("SIGNTOOL", "signtool")):
            raise RuntimeError("Set SIGNTOOL to the Windows SDK signtool.exe path.")


def sign_windows(path):
    tool = os.environ.get("SIGNTOOL", "signtool")
    subprocess.run([tool, "sign", "/sha1", setting("WINDOWS_CERT_SHA1"),
                    "/fd", "SHA256", "/tr", setting("WINDOWS_TIMESTAMP_URL"),
                    "/td", "SHA256", str(path)], check=True, timeout=180)
    subprocess.run([tool, "verify", "/pa", "/all", "/v", str(path)], check=True, timeout=60)


def notarize(path):
    subprocess.run(["xcrun", "notarytool", "submit", str(path),
                    "--keychain-profile", setting("APPLE_NOTARY_PROFILE"),
                    "--wait", "--timeout", "20m"], check=True, timeout=1260)
    subprocess.run(["xcrun", "stapler", "staple", str(path)], check=True, timeout=120)
    subprocess.run(["xcrun", "stapler", "validate", str(path)], check=True, timeout=60)


def make_dmg(app, destination, signed=False):
    import tempfile

    with tempfile.TemporaryDirectory(prefix="thesiscraft-dmg-") as folder:
        stage = Path(folder)
        subprocess.run(["ditto", str(app), str(stage / app.name)], check=True)
        (stage / "Applications").symlink_to("/Applications")
        subprocess.run(["hdiutil", "create", "-volname", "ThesisCraft", "-srcfolder", str(stage),
                        "-ov", "-format", "UDZO", str(destination)], check=True)
    if signed:
        subprocess.run(["codesign", "--sign", setting("APPLE_SIGNING_IDENTITY"),
                        "--timestamp", str(destination)], check=True)
        notarize(destination)


def notarize_app(app, archive):
    # Submit ZIP, staple the contained app, then let the caller rebuild its ZIP.
    subprocess.run(["xcrun", "notarytool", "submit", str(archive), "--keychain-profile",
                    setting("APPLE_NOTARY_PROFILE"), "--wait", "--timeout", "20m"],
                   check=True, timeout=1260)
    subprocess.run(["xcrun", "stapler", "staple", str(app)], check=True, timeout=120)
    subprocess.run(["xcrun", "stapler", "validate", str(app)], check=True, timeout=60)
    subprocess.run(["spctl", "--assess", "--type", "execute", "--verbose", str(app)], check=True)
