import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime

from runner_utils import (
    compare_versions,
    download_text,
    download_to_path,
    extract_sha256_from_text,
    extract_zip,
    extract_latest_filenames,
    is_newer_version,
    is_prerelease,
    load_json_file,
    normalize_staging_layout,
    read_version_file,
    resolve_network_asset_paths,
    sha256_file,
    validate_latest_json,
    write_version_file,
)


LEVELS = {"DEBUG": 10, "INFO": 20, "WARN": 30, "ERROR": 40}


class RunnerLogger:
    def __init__(self, log_file: str, log_level: str = "INFO"):
        self.log_file = log_file
        self.log_level = (log_level or "INFO").upper()
        if self.log_file:
            os.makedirs(os.path.dirname(self.log_file), exist_ok=True)

    def _should_log(self, level: str) -> bool:
        current = LEVELS.get(self.log_level, 20)
        incoming = LEVELS.get(level.upper(), 20)
        return incoming >= current

    def _write(self, level: str, message: str):
        if not self._should_log(level):
            return
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] [{level}] {message}"
        print(line)
        if self.log_file:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(line + "\n")

    def info(self, message: str):
        self._write("INFO", message)

    def warn(self, message: str):
        self._write("WARN", message)

    def error(self, message: str):
        self._write("ERROR", message)


def load_runner_config(path: str) -> dict:
    if not os.path.exists(path):
        raise FileNotFoundError(f"config.json nao encontrado: {path}")
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    install_dir = data.get("install_dir") or data.get("install_root")
    required = ["app_name", "network_release_dir", "network_latest_json", "exe_name", "log_file"]
    missing = [k for k in required if k not in data or not data.get(k)]
    if not install_dir:
        missing.append("install_dir")
    if missing:
        raise ValueError(f"config.json invalido, faltando: {', '.join(missing)}")
    data["install_dir"] = install_dir
    data["prefer_network"] = bool(data.get("prefer_network", True))
    data["allow_prerelease"] = bool(data.get("allow_prerelease", False))
    data["run_args"] = data.get("run_args") or []
    if not isinstance(data["run_args"], list):
        raise ValueError("run_args deve ser uma lista")
    data["log_level"] = (data.get("log_level") or "INFO").upper()
    github_repo = data.get("github_repo") or ""
    data["github_repo"] = github_repo.strip()
    return data


def fetch_latest_from_network(config: dict, logger: RunnerLogger) -> dict:
    latest_path = config["network_latest_json"]
    logger.info(f"Buscando latest.json na rede: {latest_path}")
    data = load_json_file(latest_path)
    validate_latest_json(data)
    package_filename, sha256_filename = extract_latest_filenames(data)
    data["package_filename"] = package_filename
    data["sha256_filename"] = sha256_filename
    zip_path, sha_path = resolve_network_asset_paths(data, config["network_release_dir"])
    data["_zip_path"] = zip_path
    data["_sha_path"] = sha_path
    data["_source"] = "network"
    return data


def fetch_latest_from_github(config: dict, logger: RunnerLogger) -> dict:
    repo = config["github_repo"]
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    logger.info(f"Buscando latest release no GitHub: {repo}")
    raw = download_text(url)
    data = json.loads(raw)
    tag = data.get("tag_name") or data.get("name")
    if not tag:
        raise ValueError("Release sem tag_name/name no GitHub")
    version = tag.lstrip("v").strip()
    compare_versions(version, "0.0.0")

    assets = data.get("assets") or []
    zip_asset = None
    sha_asset = None
    for asset in assets:
        name = (asset.get("name") or "").lower()
        if name.endswith(".zip") and not zip_asset:
            zip_asset = asset
        if name.endswith(".sha256") and not sha_asset:
            sha_asset = asset
        if config["app_name"].lower() in name:
            if name.endswith(".zip"):
                zip_asset = asset
            if name.endswith(".sha256"):
                sha_asset = asset

    if not zip_asset or not sha_asset:
        raise ValueError("Assets .zip/.sha256 nao encontrados no GitHub")

    return {
        "version": version,
        "package_filename": zip_asset.get("name"),
        "sha256_filename": sha_asset.get("name"),
        "zip_name": zip_asset.get("name"),
        "sha256_name": sha_asset.get("name"),
        "_zip_url": zip_asset.get("browser_download_url"),
        "_sha_url": sha_asset.get("browser_download_url"),
        "_source": "github",
    }


def validate_sha256(zip_path: str, sha_path: str, logger: RunnerLogger) -> bool:
    with open(sha_path, "r", encoding="utf-8") as f:
        sha_text = f.read()
    expected = extract_sha256_from_text(sha_text)
    actual = sha256_file(zip_path)
    if expected != actual:
        logger.error("SHA256 divergente. Instalacao abortada.")
        return False
    logger.info("SHA256 validado com sucesso.")
    return True


def _is_exe_in_use_error(error: Exception) -> bool:
    if not isinstance(error, OSError):
        return False
    winerror = getattr(error, "winerror", None)
    return winerror in (32, 5)


def _install_paths(install_dir: str) -> dict:
    app_root = os.path.join(install_dir, "app")
    return {
        "install_dir": install_dir,
        "app_root": app_root,
        "current_dir": os.path.join(app_root, "current"),
        "staging_dir": os.path.join(app_root, "staging"),
        "backup_dir": os.path.join(app_root, "backup"),
        "logs_dir": os.path.join(install_dir, "logs"),
        "version_file": os.path.join(install_dir, "version.txt"),
    }


def perform_install(config: dict, latest: dict, logger: RunnerLogger) -> str:
    paths = _install_paths(config["install_dir"])
    current_dir = paths["current_dir"]
    staging_dir = paths["staging_dir"]
    backup_dir = paths["backup_dir"]

    zip_path = latest.get("_zip_path")
    sha_path = latest.get("_sha_path")

    if not zip_path or not sha_path:
        raise ValueError("Caminhos do zip/sha nao resolvidos")

    logger.info(f"Extraindo zip para staging: {staging_dir}")
    extract_zip(zip_path, staging_dir)
    normalize_staging_layout(staging_dir, config["exe_name"])

    backup_made = False
    try:
        if os.path.exists(backup_dir):
            shutil.rmtree(backup_dir, ignore_errors=True)
        if os.path.exists(current_dir):
            logger.info("Movendo current para backup...")
            try:
                os.replace(current_dir, backup_dir)
            except Exception as e:
                if _is_exe_in_use_error(e):
                    shutil.rmtree(staging_dir, ignore_errors=True)
                    return "deferred"
                raise
            backup_made = True

        logger.info("Movendo staging para current...")
        try:
            os.replace(staging_dir, current_dir)
        except Exception as e:
            if _is_exe_in_use_error(e):
                shutil.rmtree(staging_dir, ignore_errors=True)
                if backup_made and os.path.exists(backup_dir) and not os.path.exists(current_dir):
                    try:
                        logger.warn("Rollback por exe em uso...")
                        os.replace(backup_dir, current_dir)
                        logger.warn("Rollback concluido.")
                    except Exception as rollback_error:
                        logger.error(f"Rollback falhou: {rollback_error}")
                return "deferred"
            raise
        write_version_file(paths["version_file"], latest["version"])
        logger.info("Instalacao concluida.")
        return "ok"
    except Exception as e:
        logger.error(f"Falha ao instalar: {e}")
        if backup_made and os.path.exists(backup_dir) and not os.path.exists(current_dir):
            try:
                logger.warn("Rollback iniciado...")
                os.replace(backup_dir, current_dir)
                logger.warn("Rollback concluido.")
            except Exception as rollback_error:
                logger.error(f"Rollback falhou: {rollback_error}")
        return "failed"


def download_github_assets(latest: dict, logger: RunnerLogger) -> tuple[str, str]:
    tmp_dir = tempfile.mkdtemp(prefix="metaxg_download_")
    package_filename, sha256_filename = extract_latest_filenames(latest)
    zip_path = os.path.join(tmp_dir, package_filename)
    sha_path = os.path.join(tmp_dir, sha256_filename)
    logger.info(f"Baixando zip do GitHub: {latest['_zip_url']}")
    download_to_path(latest["_zip_url"], zip_path)
    logger.info(f"Baixando sha256 do GitHub: {latest['_sha_url']}")
    download_to_path(latest["_sha_url"], sha_path)
    latest["_zip_path"] = zip_path
    latest["_sha_path"] = sha_path
    return zip_path, sha_path


def run_app(config: dict, app_args: list[str], logger: RunnerLogger):
    paths = _install_paths(config["install_dir"])
    current_dir = paths["current_dir"]
    exe_path = os.path.join(current_dir, config["exe_name"])
    if not os.path.exists(exe_path):
        logger.error(f"Exe nao encontrado: {exe_path}")
        return 1
    logger.info(f"Executando: {exe_path}")
    try:
        completed = subprocess.run([exe_path, *app_args], cwd=current_dir)
        return completed.returncode
    except Exception as e:
        logger.error(f"Falha ao executar exe: {e}")
        return 1


def parse_args():
    parser = argparse.ArgumentParser(description="Runner MetaXg")
    parser.add_argument("--config", default="config.json", help="Caminho do config.json do runner")
    return parser.parse_known_args()


def build_source_order(prefer_network: bool, github_enabled: bool) -> list[str]:
    if prefer_network:
        order = ["network"]
        if github_enabled:
            order.append("github")
        return order
    order = []
    if github_enabled:
        order.append("github")
    order.append("network")
    return order


def main():
    args, app_args = parse_args()
    config_path = args.config
    if not os.path.isabs(config_path):
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), config_path)

    config = load_runner_config(config_path)
    logger = RunnerLogger(config["log_file"], log_level=config["log_level"])
    logger.info("===== RUNNER START =====")

    install_dir = config["install_dir"]
    paths = _install_paths(install_dir)
    os.makedirs(paths["install_dir"], exist_ok=True)
    os.makedirs(paths["app_root"], exist_ok=True)
    os.makedirs(paths["logs_dir"], exist_ok=True)
    current_version = read_version_file(paths["version_file"])
    logger.info(f"Versao instalada: {current_version}")

    latest = None
    github_enabled = bool(config.get("github_repo"))
    for source in build_source_order(config["prefer_network"], github_enabled):
        try:
            if source == "network":
                latest = fetch_latest_from_network(config, logger)
                logger.info(f"Latest (rede): {latest['version']}")
            else:
                latest = fetch_latest_from_github(config, logger)
                logger.info(f"Latest (GitHub): {latest['version']}")
            if latest and is_prerelease(latest["version"]) and not config["allow_prerelease"]:
                logger.warn("Prerelease detectado e allow_prerelease=false. Ignorando.")
                latest = None
                continue
            break
        except Exception as e:
            logger.warn(f"Falha ao ler latest ({source}): {e}")

    if latest and is_newer_version(latest["version"], current_version):
        logger.info("Atualizacao disponivel. Iniciando update...")
        try:
            if latest.get("_source") == "github":
                download_github_assets(latest, logger)

            if not validate_sha256(latest["_zip_path"], latest["_sha_path"], logger):
                logger.warn("SHA invalido. Executando versao atual.")
            else:
                result = perform_install(config, latest, logger)
                if result == "deferred":
                    logger.warn("exe em uso, update adiado")
                elif result != "ok":
                    logger.warn("Instalacao falhou. Executando versao atual.")
        except Exception as e:
            logger.error(f"Erro no processo de update: {e}")
            logger.warn("Executando versao atual.")
    else:
        if latest:
            logger.info("Nenhuma atualizacao necessaria.")
        else:
            logger.warn("Sem latest disponivel. Executando versao atual.")

    final_args = list(config.get("run_args") or []) + list(app_args)
    code = run_app(config, final_args, logger)
    logger.info(f"===== RUNNER END (exit={code}) =====")
    sys.exit(code)


if __name__ == "__main__":
    main()
