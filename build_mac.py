# md2word — macOS PyInstaller build & packaging script
# Usage:  python3 build_mac.py

from __future__ import annotations

import os
import platform
import plistlib
import re
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
ASSETS = ROOT / "assets"
NAME = "md2word"
BUNDLE_ID = "com.wanxiao2018.md2word"


def get_version() -> str:
    pyproject = ROOT / "pyproject.toml"
    if pyproject.is_file():
        m = re.search(r'version\s*=\s*"([^"]+)"', pyproject.read_text(encoding="utf-8"))
        if m:
            return m.group(1)
    return "1.2.0"


def ensure_icns() -> Path:
    icns_path = ASSETS / "app.icns"
    if icns_path.is_file():
        return icns_path

    png_path = ASSETS / "app.png"
    if not png_path.is_file():
        print("Warning: Neither app.icns nor app.png found in assets/")
        return icns_path

    print("Generating assets/app.icns from assets/app.png...")
    iconset = ASSETS / "app.iconset"
    if iconset.exists():
        shutil.rmtree(iconset)
    iconset.mkdir(parents=True, exist_ok=True)

    sizes = [
        (16, "16x16"),
        (32, "16x16@2x"),
        (32, "32x32"),
        (64, "32x32@2x"),
        (128, "128x128"),
        (256, "128x128@2x"),
        (256, "256x256"),
    ]
    try:
        for sz, tag in sizes:
            out_file = iconset / f"icon_{tag}.png"
            subprocess.run(
                ["sips", "-z", str(sz), str(sz), str(png_path), "--out", str(out_file)],
                check=True,
                capture_output=True,
            )
        subprocess.run(
            ["iconutil", "-c", "icns", str(iconset), "-o", str(icns_path)],
            check=True,
            capture_output=True,
        )
        print(f"Generated {icns_path}")
    except Exception as exc:  # noqa: BLE001
        print(f"Warning: Failed to create app.icns via iconutil: {exc}")
    finally:
        if iconset.exists():
            shutil.rmtree(iconset)

    return icns_path


def build_app(version: str) -> Path:
    icns = ensure_icns()
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        f"--name={NAME}",
        f"--paths={ROOT}",
        f"--osx-bundle-identifier={BUNDLE_ID}",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        "--hidden-import=lxml.etree",
        "--exclude-module=pygame",
        "--exclude-module=matplotlib",
        "--exclude-module=numpy",
        "--exclude-module=torch",
        "--exclude-module=pandas",
        "--exclude-module=scipy",
        "--exclude-module=cv2",
        str(ROOT / "main.py"),
    ]

    sep = ":"
    if icns.is_file():
        cmd.insert(-1, f"--icon={icns}")
        cmd.insert(-1, f"--add-data={icns}{sep}assets")

    png = ASSETS / "app.png"
    if png.is_file():
        cmd.insert(-1, f"--add-data={png}{sep}assets")

    hpng = ASSETS / "app-32.png"
    if hpng.is_file():
        cmd.insert(-1, f"--add-data={hpng}{sep}assets")

    ico = ASSETS / "app.ico"
    if ico.is_file():
        cmd.insert(-1, f"--add-data={ico}{sep}assets")

    print("\nRunning PyInstaller...")
    print(" ".join(cmd))
    res = subprocess.run(cmd, cwd=str(ROOT))
    if res.returncode != 0:
        raise RuntimeError(f"PyInstaller failed with code {res.returncode}")

    app_path = DIST / f"{NAME}.app"
    if not app_path.is_dir():
        raise FileNotFoundError(f"Expected app bundle not found: {app_path}")

    # Refine Info.plist
    plist_path = app_path / "Contents" / "Info.plist"
    if plist_path.is_file():
        try:
            with open(plist_path, "rb") as f:
                pl = plistlib.load(f)
            pl["CFBundleDisplayName"] = NAME
            pl["CFBundleName"] = NAME
            pl["CFBundleIdentifier"] = BUNDLE_ID
            pl["CFBundleShortVersionString"] = version
            pl["CFBundleVersion"] = version
            pl["NSHighResolutionCapable"] = True
            pl["LSMinimumSystemVersion"] = "10.15"
            with open(plist_path, "wb") as f:
                plistlib.dump(pl, f)
            print(f"Updated Info.plist with version {version}")
        except Exception as exc:  # noqa: BLE001
            print(f"Warning: Could not update Info.plist: {exc}")

    # Ad-hoc codesign
    try:
        subprocess.run(
            ["codesign", "--force", "--deep", "-s", "-", str(app_path)],
            check=True,
            capture_output=True,
        )
        print("Ad-hoc code signing completed.")
    except Exception as exc:  # noqa: BLE001
        print(f"Warning: codesign failed: {exc}")

    return app_path


def package_zip(app_path: Path, arch: str) -> list[Path]:
    zip_arch = DIST / f"{NAME}-macos-{arch}.zip"
    zip_gen = DIST / f"{NAME}-macos.zip"

    print(f"\nCreating ZIP archive: {zip_arch.name}...")
    subprocess.run(
        ["ditto", "-c", "-k", "--sequesterRsrc", "--keepParent", str(app_path), str(zip_arch)],
        check=True,
    )
    shutil.copy2(zip_arch, zip_gen)
    return [zip_arch, zip_gen]


def package_dmg(app_path: Path, arch: str) -> list[Path]:
    dmg_arch = DIST / f"{NAME}-macos-{arch}.dmg"
    dmg_gen = DIST / f"{NAME}-macos.dmg"

    print(f"\nCreating DMG image: {dmg_arch.name}...")
    staging = DIST / "_dmg_staging"
    if staging.exists():
        shutil.rmtree(staging)
    staging.mkdir(parents=True, exist_ok=True)

    try:
        shutil.copytree(app_path, staging / f"{NAME}.app", symlinks=True)
        # Create Applications shortcut link
        os.symlink("/Applications", staging / "Applications")

        if dmg_arch.exists():
            dmg_arch.unlink()

        subprocess.run(
            [
                "hdiutil",
                "create",
                "-volname",
                NAME,
                "-srcfolder",
                str(staging),
                "-ov",
                "-format",
                "UDZO",
                str(dmg_arch),
            ],
            check=True,
            capture_output=True,
        )
        shutil.copy2(dmg_arch, dmg_gen)
        return [dmg_arch, dmg_gen]
    finally:
        if staging.exists():
            shutil.rmtree(staging)


def main() -> int:
    if sys.platform != "darwin":
        print("Error: build_mac.py must be run on macOS.", file=sys.stderr)
        return 1

    version = get_version()
    arch = platform.machine().lower()
    if arch in ("x86_64", "amd64"):
        arch = "x64"

    print(f"=== Building {NAME} v{version} for macOS ({arch}) ===")
    app_path = build_app(version)

    zips = package_zip(app_path, arch)
    dmgs = package_dmg(app_path, arch)

    print("\n=== Build Completed Successfully ===")
    print(f"App Bundle: {app_path}")
    for p in dmgs + zips:
        mb = p.stat().st_size / (1024 * 1024)
        print(f"  - {p.name}: {mb:.1f} MB ({p})")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
