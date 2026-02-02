import json
import os
import zipfile

import pytest

import runner
from runner import build_source_order, load_runner_config, perform_install
from runner_utils import (
    compare_versions,
    extract_latest_filenames,
    extract_sha256_from_text,
    is_prerelease,
    is_newer_version,
    resolve_network_asset_paths,
    sha256_file,
    validate_latest_json,
)


def test_compare_versions():
    assert compare_versions("1.2.3", "1.2.3") == 0
    assert compare_versions("1.2.4", "1.2.3") == 1
    assert compare_versions("2.0.0", "2.1.0") == -1
    assert compare_versions("1.2.3", "1.2.3-beta") == 1


def test_is_newer_version():
    assert is_newer_version("1.2.4", "1.2.3") is True
    assert is_newer_version("1.2.3", "1.2.3") is False
    assert is_prerelease("1.2.3-beta") is True


def test_validate_latest_json():
    data = {"version": "1.2.3", "package_filename": "MetaXg_1.2.3.zip", "sha256_filename": "MetaXg_1.2.3.sha256"}
    validate_latest_json(data)
    with pytest.raises(ValueError):
        validate_latest_json({"version": "1.2.3"})


def test_extract_sha256_from_text():
    sha = "a" * 64
    text = f"{sha}  MetaXg_1.2.3.zip"
    assert extract_sha256_from_text(text) == sha


def test_sha256_file(tmp_path):
    file_path = tmp_path / "data.bin"
    file_path.write_bytes(b"abc")
    assert sha256_file(str(file_path)) == "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"


def test_resolve_network_asset_paths(tmp_path):
    latest = {"package_filename": "MetaXg_1.2.3.zip", "sha256_filename": "MetaXg_1.2.3.sha256"}
    zip_path, sha_path = resolve_network_asset_paths(latest, str(tmp_path))
    assert zip_path.endswith("MetaXg_1.2.3.zip")
    assert sha_path.endswith("MetaXg_1.2.3.sha256")


def test_extract_latest_filenames_legacy():
    latest = {"zip_name": "MetaXg_1.2.3.zip", "sha256_name": "MetaXg_1.2.3.sha256"}
    pkg, sha = extract_latest_filenames(latest)
    assert pkg.endswith(".zip")
    assert sha.endswith(".sha256")


def test_perform_install(tmp_path):
    install_root = tmp_path / "install"
    current_dir = install_root / "app" / "current"
    current_dir.mkdir(parents=True)
    (current_dir / "old.txt").write_text("old", encoding="utf-8")

    staging_dir = tmp_path / "pkg"
    staging_dir.mkdir(parents=True)
    (staging_dir / "MetaXg.exe").write_text("exe", encoding="utf-8")
    (staging_dir / "config.json").write_text(json.dumps({"ok": True}), encoding="utf-8")

    zip_path = tmp_path / "MetaXg_1.2.3.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        for item in staging_dir.iterdir():
            zf.write(item, arcname=item.name)

    sha_path = tmp_path / "MetaXg_1.2.3.sha256"
    sha_path.write_text(f"{sha256_file(str(zip_path))}  MetaXg_1.2.3.zip", encoding="utf-8")

    config = {"install_dir": str(install_root), "exe_name": "MetaXg.exe"}
    latest = {"version": "1.2.3", "_zip_path": str(zip_path), "_sha_path": str(sha_path)}

    class DummyLogger:
        def info(self, _):
            pass
        def warn(self, _):
            pass
        def error(self, _):
            pass

    result = perform_install(config, latest, DummyLogger())
    assert result == "ok"
    assert os.path.exists(install_root / "app" / "current" / "MetaXg.exe")
    assert os.path.exists(install_root / "app" / "backup" / "old.txt")
    assert (install_root / "version.txt").read_text(encoding="utf-8") == "1.2.3"


def test_build_source_order():
    assert build_source_order(True, False) == ["network"]
    assert build_source_order(False, False) == ["network"]
    assert build_source_order(False, True) == ["github", "network"]


def test_load_runner_config_github_optional(tmp_path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "app_name": "MetaXg",
                "install_dir": "C:\\MetaXg",
                "network_release_dir": "P:\\ProcessoMetaX\\releases",
                "network_latest_json": "P:\\ProcessoMetaX\\releases\\latest.json",
                "exe_name": "MetaXg.exe",
                "log_file": "C:\\MetaXg\\logs\\metax_last_run.log",
            }
        ),
        encoding="utf-8",
    )
    cfg = load_runner_config(str(config_path))
    assert cfg["github_repo"] == ""


def test_perform_install_deferred_when_exe_in_use(tmp_path, monkeypatch):
    install_root = tmp_path / "install"
    current_dir = install_root / "app" / "current"
    current_dir.mkdir(parents=True)
    (current_dir / "MetaXg.exe").write_text("exe", encoding="utf-8")

    staging_dir = tmp_path / "pkg"
    staging_dir.mkdir(parents=True)
    (staging_dir / "MetaXg.exe").write_text("exe", encoding="utf-8")

    zip_path = tmp_path / "MetaXg_1.2.3.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        for item in staging_dir.iterdir():
            zf.write(item, arcname=item.name)

    sha_path = tmp_path / "MetaXg_1.2.3.sha256"
    sha_path.write_text(f"{sha256_file(str(zip_path))}  MetaXg_1.2.3.zip", encoding="utf-8")

    config = {"install_dir": str(install_root), "exe_name": "MetaXg.exe"}
    latest = {"version": "1.2.3", "_zip_path": str(zip_path), "_sha_path": str(sha_path)}

    original_replace = runner.os.replace

    def fake_replace(src, dst):
        if src == str(current_dir):
            err = OSError("sharing violation")
            err.winerror = 32
            raise err
        return original_replace(src, dst)

    monkeypatch.setattr(runner.os, "replace", fake_replace)

    class DummyLogger:
        def info(self, _):
            pass
        def warn(self, _):
            pass
        def error(self, _):
            pass

    result = perform_install(config, latest, DummyLogger())
    assert result == "deferred"
    assert os.path.exists(current_dir)
