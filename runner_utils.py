import json
import os
import re
import shutil
import tempfile
import zipfile
from hashlib import sha256
from typing import Tuple
from urllib.request import urlopen, Request


SEMVER_BASE_RE = re.compile(r"^(?P<major>0|[1-9]\d*)\.(?P<minor>0|[1-9]\d*)\.(?P<patch>0|[1-9]\d*)$")


def split_version(version: str) -> Tuple[int, int, int, str]:
    if not version:
        raise ValueError("version vazio")
    version = version.strip()
    if version.startswith("v"):
        version = version[1:]
    base, prerelease = (version.split("-", 1) + [""])[:2]
    match = SEMVER_BASE_RE.match(base)
    if not match:
        raise ValueError(f"versao invalida: {version}")
    return int(match.group("major")), int(match.group("minor")), int(match.group("patch")), prerelease


def is_prerelease(version: str) -> bool:
    return "-" in (version or "")


def parse_semver(version: str) -> Tuple[int, int, int]:
    major, minor, patch, _ = split_version(version)
    return major, minor, patch


def _compare_prerelease(a: str, b: str) -> int:
    if not a and not b:
        return 0
    if not a and b:
        return 1
    if a and not b:
        return -1
    a_ids = a.split(".")
    b_ids = b.split(".")
    for i in range(max(len(a_ids), len(b_ids))):
        if i >= len(a_ids):
            return -1
        if i >= len(b_ids):
            return 1
        ai = a_ids[i]
        bi = b_ids[i]
        ai_is_num = ai.isdigit()
        bi_is_num = bi.isdigit()
        if ai_is_num and bi_is_num:
            ai_num = int(ai)
            bi_num = int(bi)
            if ai_num != bi_num:
                return 1 if ai_num > bi_num else -1
        elif ai_is_num and not bi_is_num:
            return -1
        elif not ai_is_num and bi_is_num:
            return 1
        else:
            if ai != bi:
                return 1 if ai > bi else -1
    return 0


def compare_versions(a: str, b: str) -> int:
    amajor, aminor, apatch, apre = split_version(a)
    bmajor, bminor, bpatch, bpre = split_version(b)
    ta = (amajor, aminor, apatch)
    tb = (bmajor, bminor, bpatch)
    if ta != tb:
        return 1 if ta > tb else -1
    return _compare_prerelease(apre, bpre)


def is_newer_version(latest: str, current: str) -> bool:
    return compare_versions(latest, current) > 0


def read_version_file(path: str) -> str:
    if not os.path.exists(path):
        return "0.0.0"
    with open(path, "r", encoding="utf-8") as f:
        return f.read().strip() or "0.0.0"


def write_version_file(path: str, version: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(version)


def load_json_file(path: str) -> dict:
    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def extract_latest_filenames(data: dict) -> tuple[str, str]:
    package_filename = data.get("package_filename") or data.get("zip_name")
    sha256_filename = data.get("sha256_filename") or data.get("sha256_name")
    if not package_filename or not sha256_filename:
        raise ValueError("latest.json invalido, faltando package_filename/sha256_filename")
    return package_filename, sha256_filename


def validate_latest_json(data: dict):
    if not data.get("version"):
        raise ValueError("latest.json invalido, faltando version")
    extract_latest_filenames(data)


def resolve_network_asset_paths(latest_json: dict, network_release_dir: str) -> tuple[str, str]:
    package_filename, sha256_filename = extract_latest_filenames(latest_json)
    zip_path = package_filename if os.path.isabs(package_filename) else os.path.join(network_release_dir, package_filename)
    sha_path = sha256_filename if os.path.isabs(sha256_filename) else os.path.join(network_release_dir, sha256_filename)
    return zip_path, sha_path


def extract_sha256_from_text(text: str) -> str:
    match = re.search(r"\b[a-fA-F0-9]{64}\b", text or "")
    if not match:
        raise ValueError("SHA256 nao encontrado no arquivo .sha256")
    return match.group(0).lower()


def sha256_file(path: str) -> str:
    h = sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def download_to_path(url: str, dest_path: str, timeout_sec: int = 30):
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    req = Request(url, headers={"User-Agent": "MetaXgRunner/1.0"})
    with urlopen(req, timeout=timeout_sec) as resp, open(dest_path, "wb") as f:
        f.write(resp.read())


def download_text(url: str, timeout_sec: int = 30) -> str:
    req = Request(url, headers={"User-Agent": "MetaXgRunner/1.0"})
    with urlopen(req, timeout=timeout_sec) as resp:
        return resp.read().decode("utf-8", errors="replace")


def clean_dir(path: str):
    if os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)
    os.makedirs(path, exist_ok=True)


def extract_zip(zip_path: str, dest_dir: str):
    clean_dir(dest_dir)
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(dest_dir)


def normalize_staging_layout(staging_dir: str, exe_name: str):
    exe_path = os.path.join(staging_dir, exe_name)
    if os.path.exists(exe_path):
        return

    entries = [e for e in os.listdir(staging_dir) if not e.startswith(".")]
    if len(entries) == 1:
        candidate = os.path.join(staging_dir, entries[0])
        if os.path.isdir(candidate) and os.path.exists(os.path.join(candidate, exe_name)):
            tmp_dir = tempfile.mkdtemp(prefix="metaxg_stage_")
            try:
                for item in os.listdir(candidate):
                    shutil.move(os.path.join(candidate, item), tmp_dir)
                clean_dir(staging_dir)
                for item in os.listdir(tmp_dir):
                    shutil.move(os.path.join(tmp_dir, item), staging_dir)
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            return

    raise FileNotFoundError(f"Exe {exe_name} nao encontrado no staging")


def safe_replace(src: str, dst: str):
    if os.path.exists(dst):
        shutil.rmtree(dst, ignore_errors=True)
    os.replace(src, dst)
