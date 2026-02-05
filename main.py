import argparse
import os
import re
import shutil
import unicodedata
import uuid
from datetime import datetime

import pyodbc
from custom_logger import logger
from auditoria_excel import log_run
from output_manager import OutputManager, KIND_JSON
from outcomes import (
    OUTCOME_FAILED_ACTION,
    OUTCOME_FAILED_VERIFICATION,
    OUTCOME_SAVED_NOT_VERIFIED,
    OUTCOME_VERIFIED_SUCCESS,
    OUTCOME_SKIPPED_ALREADY_EXISTS,
    OUTCOME_SKIPPED_DRY_RUN,
    OUTCOME_SKIPPED_EMAIL_DISABLED,
    compute_totals,
)
from rpa_metax import iniciar_sessao, cadastrar_funcionario, obter_todos_rascunhos, verificar_cadastro
from sharepoint import baixar_fotos_em_lote

from config import (
    DB_DRIVER, DB_SERVER, DB_NAME, DB_USER, DB_PASSWORD, DIAS_RETROATIVOS,
    ROOT_DIR, PUBLIC_BASE_DIR, PUBLIC_INPUTS_DIR, OBJECT_NAME,
    PUBLIC_CODE_DIR, PUBLIC_PROCESSADOS_DIR, PUBLIC_ERROS_DIR,
    PUBLIC_LOGS_DIR, PUBLIC_RELATORIOS_DIR, PUBLIC_JSON_DIR, PUBLIC_RELEASES_DIR,
    FOTOS_EM_PROCESSAMENTO_DIR, FOTOS_PROCESSADOS_DIR, FOTOS_ERROS_DIR, FOTOS_BUSCA_DIRS,
    METAX_CONTRATO_MECANICA_VALUE, METAX_CONTRATO_MECANICA_LABEL,
    METAX_CONTRATO_ELETROMECANICA_VALUE, METAX_CONTRATO_ELETROMECANICA_LABEL
)


def obter_conexao() -> pyodbc.Connection:
    """Estabelece conexao com o banco de dados SQL Server."""
    drivers_alternativos = [
        DB_DRIVER,
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 18 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "SQL Server",
        "SQL Server Native Client 11.0",
    ]

    drivers_disponiveis = [d for d in pyodbc.drivers()]

    ultimo_erro = None
    for driver in drivers_alternativos:
        if driver in drivers_disponiveis:
            try:
                logger.info(f"Tentando conectar com driver: {driver}", details={"driver": driver})
                conexao = pyodbc.connect(
                    f"DRIVER={{{driver}}};"
                    f"SERVER={DB_SERVER};"
                    f"DATABASE={DB_NAME};"
                    f"UID={DB_USER};"
                    f"PWD={DB_PASSWORD}"
                )
                logger.info(f"Conexao estabelecida com sucesso usando driver: {driver}", details={"driver": driver})
                return conexao
            except Exception as e:
                ultimo_erro = e
                logger.warn(f"Falha ao conectar com driver {driver}: {e}", details={"driver": driver, "erro": str(e)})
                continue

    if ultimo_erro:
        raise ConnectionError(f"Nao foi possivel conectar ao banco de dados. Ultimo erro: {ultimo_erro}")

    raise ConnectionError(
        f"Nenhum driver ODBC para SQL Server encontrado. "
        f"Drivers disponiveis: {', '.join(drivers_disponiveis)}. "
        f"Por favor, instale um driver ODBC para SQL Server."
    )


def buscar_funcionarios_para_cadastro(data_admissao: str = None, filtro_nomes: list[str] = None) -> list[dict]:
    """
    Busca funcionarios.
    - Se filtro_nomes for fornecido, busca APENAS esses nomes (ignora data).
    - Se nao, busca por data de admissao (e dias retroativos).
    """
    if not data_admissao:
        data_admissao = datetime.now().strftime("%Y-%m-%d")

    datas_vars = f"""
        DECLARE @DataFim DATE = '{data_admissao}';
        DECLARE @DataInicio DATE = DATEADD(DAY, -{DIAS_RETROATIVOS}, @DataFim);
    """

    params = []
    if filtro_nomes and len(filtro_nomes) > 0:
        logger.info(f"MODO FILTRO ATIVADO: Buscando {len(filtro_nomes)} funcionario(s) especifico(s).")
        placeholders = ", ".join(["?"] * len(filtro_nomes))
        where_clause = f"UPPER(P.NOME) IN ({placeholders})"
        params = [nome.strip().upper() for nome in filtro_nomes]
    else:
        if DIAS_RETROATIVOS > 0:
            logger.info(
                f"Buscando funcionarios com admissao entre {data_admissao} e {DIAS_RETROATIVOS} dia(s) antes.",
                details={"data_ref": data_admissao, "retroativos": DIAS_RETROATIVOS},
            )
        else:
            logger.info(
                f"Buscando funcionarios com admissao em: {data_admissao}",
                details={"data_admissao": data_admissao},
            )
        where_clause = "CAST(F.DATAADMISSAO AS DATE) BETWEEN @DataInicio AND @DataFim"

    sql = f"""
        {datas_vars}

        SELECT 
            P.NOME,
            P.CPF,
            F.CHAPA,
            P.SEXO,
            P.NATURALIDADE,
            P.GRAUINSTRUCAO,
            P.ESTADOCIVIL,
            P.ESTADONATAL,
            P.DTNASCIMENTO,
            P.EMAIL,
            P.TELEFONE1,

            COALESCE(PAI.NOME, 'NAO INFORMADO') AS NOME_PAI,
            COALESCE(MAE.NOME, 'NAO INFORMADO') AS NOME_MAE,

            P.ORGEMISSORIDENT,
            P.UFCARTIDENT,
            P.CARTIDENTIDADE,
            P.DTEMISSAOIDENT,
            P.CARTEIRATRAB,
            P.SERIECARTTRAB,
            P.UFCARTTRAB,
            P.DTCARTTRAB,
            P.TITULOELEITOR,
            P.ZONATITELEITOR,
            P.SECAOTITELEITOR,
            P.ESTELEIT,

            P.CEP,
            P.ESTADO,
            P.BAIRRO,
            P.RUA,
            P.NUMERO,

            F.PISPASEP,
            F.DATAADMISSAO,
            F.SALARIO,
            F.CODFUNCAO,
            SEC.NROCENCUSTOCONT AS CENTRO_CUSTO,

            COALESCE(
                UPPER(LTRIM(RTRIM(FU.NOME))),
                'NAO INFORMADO'
            ) AS DESCRICAO_CARGO,

            TRY_CAST(
                CAST(
                    '<i>' + REPLACE(F.CODSECAO, '.', '</i><i>') + '</i>'
                    AS XML
                ).value('/i[2]', 'varchar(10)')
                AS INT
            ) AS NUMERO_OBRA

        FROM PFUNC F
        INNER JOIN PPESSOA P
            ON P.CODIGO = F.CODPESSOA

        LEFT JOIN PFUNCAO FU
            ON FU.CODCOLIGADA = F.CODCOLIGADA
        AND CAST(FU.CODIGO AS VARCHAR(20)) = F.CODFUNCAO

        LEFT JOIN PSECAO SEC
            ON SEC.CODCOLIGADA = F.CODCOLIGADA
        AND SEC.CODIGO = F.CODSECAO

        OUTER APPLY (
            SELECT TOP 1 D.NOME
            FROM PFDEPEND D
            WHERE D.CODCOLIGADA = F.CODCOLIGADA
            AND CAST(D.CHAPA AS VARCHAR(20)) = F.CHAPA
            AND CAST(D.GRAUPARENTESCO AS VARCHAR(5)) = '6'
            ORDER BY D.NOME
        ) PAI

        OUTER APPLY (
            SELECT TOP 1 D.NOME
            FROM PFDEPEND D
            WHERE D.CODCOLIGADA = F.CODCOLIGADA
            AND CAST(D.CHAPA AS VARCHAR(20)) = F.CHAPA
            AND CAST(D.GRAUPARENTESCO AS VARCHAR(5)) = '7'
            ORDER BY D.NOME
        ) MAE

        WHERE
            {where_clause}
            AND TRY_CAST(
                CAST(
                    '<i>' + REPLACE(F.CODSECAO, '.', '</i><i>') + '</i>'
                    AS XML
                ).value('/i[2]', 'varchar(10)')
                AS INT
            ) = 125
            AND P.NOME NOT IN (
                'WELISSON SANTOS SANTANA',
                'DANIELE APARECIDA RODRIGUES',
                'MARIANA CRISTINA DOS SANTOS',
                'CASSIO ROBERTO DA SILVA',
                'RAFAEL PEREIRA DA SILVA',
                'ISAIAS SOUSA LISBOA',
                'JARDELINO PEREIRA DA COSTA',
                'RAIMUNDO FRAZAO DOS SANTOS', 
                'FAUZE CELIS RODRIGUES COSTA',
                'WAGNER JUNIO DE MOURA',
                'VLADIMIR MOREIRA DA SILVA',
                'NAILTON ISAIAS MARQUES',
                'WALDEILSON PEREIRA DA SILVA',
                'WELISSON SANTOS SANTANA',
                'FRANCINALDO MARTINS SANTOS',
                'ADAILTON DE JESUS DOS SANTOS',
                'ANTONIO DOS SANTOS SILVA',
                'ADRIANO SILVA SANTOS',
                'FRANCILDO DOS SANTOS SOBRINHO',
                'JOSE ANTONIO RODRIGUES MARTINS',
                'VERILTON DOS SANTOS',
                'MANOEL JOAO PIRES SOARES',
                'ANTONIO ALVES DA SILVA FILHO',
                'EZEQUIEL DE JESUS CERQUEIRA',
                'MANOEL WALACE MACIEL SOARES',
                'EDSON PAULO MADEIRA',
                'JAIR BARROS BRANDAO',
                'ELISVELTON DA SILVA LOBATO',
                'JARDEL MENDES BRAGA',
                'HAMILSON ALVES DE MELO',
                'RAFAEL BATALHA SOUSA',
                'MARINHO DE SOUSA MACEDO',
                'ROBENILSON SANTOS CAMARA',
                'FRANCISCO VALDEVAN PAIXAO',
                'GEVANILSON CARDOSO DOS SANTOS',
                'GENILSON SILVA CANTANHEDE',
                'DENIS AUGUSTO DA ROCHA CONCEICAO',
                'EDINALDO MONTEIRO LOBATO',
                'GILSON CLEITON JOAQUIM DE MOURA',
                'MARCIO CLEITON VIEIRA DE MESQUITA',
                'MACIEL MIGUEL EVANGELISTA',
                'GILSON BARRETO SANTOS',
                'JOAO CARLOS NOVAIS SILVA',
                'ESEQUIAS FERREIRA FERNANDES',
                'JOCILEY PINHEIRO BAIA',
                'ANTONIEL MARIA ALMEIDA RIBEIRO',
                'VITOR AUGUSTO DA SILVA',
                'ERIVAN LIMA VINAGRE',
                'JOAO BATISTA SOUSA SILVA',
                'JOCICLEY DA SILVA FARIAS',
                'JOELITON ROCHA REIS',
                'DANILO FERREIRA DA SILVA',
                'PEDRO SANTOS DA SILVA',
                'MARCOS MELO MONTEIRO',
                'CARLOS EMERSON GOES OLIVA SANTOS',
                'RAIMUNDO NONATO GOMES DA SILVA',
                'ANDERSON CLAYTON PRAZERES BARRETO'
            )

        ORDER BY F.DATAADMISSAO ASC;
    """

    logger.info("Executing SQL Query", details={"sql": sql})
    with obter_conexao() as conn:
        cursor = conn.cursor()
        if params:
            cursor.execute(sql, params)
        else:
            cursor.execute(sql)
        colunas = [c[0] for c in cursor.description]
        return [dict(zip(colunas, row)) for row in cursor.fetchall()]


from notification import enviar_relatorio_email
from reporting import (
    gerar_relatorio_txt,
    gerar_relatorio_json,
    gerar_resumo_execucao_md,
    gerar_diagnostico_ultima_execucao,
)


def _normalizar_nome(raw: str) -> str:
    if not raw:
        return ""
    texto = raw.strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join([c for c in texto if not unicodedata.combining(c)])
    texto = " ".join(texto.split())
    return texto.upper()


def _extrair_bloco_centro_custo(centro_custo: str) -> str | None:
    if not centro_custo:
        return None
    texto = str(centro_custo).strip()
    partes = [p for p in re.split(r"[^\d]+", texto) if p]
    if len(partes) < 2:
        return None
    bloco = None
    # Preferencia: NROCENCUSTOCONT (ex: 125.01.004) -> bloco = partes[1]
    if len(partes) >= 3 and len(partes[1]) == 2:
        bloco = partes[1]
    # Heuristica para CODSECAO longo (ex: 36.125.001.01.01.01.01.2.004)
    elif len(partes) >= 4 and len(partes[2]) == 3:
        bloco = partes[3]
    else:
        bloco = partes[1]
    if len(bloco) == 1:
        bloco = bloco.zfill(2)
    return bloco


def _classificar_contrato_por_centro_custo(centro_custo: str) -> str:
    bloco = _extrair_bloco_centro_custo(centro_custo)
    if bloco == "01":
        return "MECANICA"
    if bloco == "02":
        return "ELETROMECANICA"
    return "DESCONHECIDO"


def _resolver_contrato_config(chave: str) -> tuple[str | None, str | None]:
    if chave == "MECANICA":
        return METAX_CONTRATO_MECANICA_VALUE, METAX_CONTRATO_MECANICA_LABEL
    if chave == "ELETROMECANICA":
        return METAX_CONTRATO_ELETROMECANICA_VALUE, METAX_CONTRATO_ELETROMECANICA_LABEL
    return None, None


def _criar_registro_base(nome: str, cpf_limpo: str, pessoa_started_at: str) -> dict:
    return {
        "nome": nome,
        "cpf": cpf_limpo,
        "attempted": False,
        "action_saved": False,
        "verified": False,
        "status_final": "FAILED",
        "outcome": OUTCOME_FAILED_ACTION,
        "errors": {
            "action_error": "",
            "verification_error": "",
        },
        "timestamps": {
            "started_at": pessoa_started_at,
            "saved_at": None,
            "verified_at": None,
        },
        "no_photo": False,
    }


def carregar_lista_nomes_txt(path: str) -> list[str]:
    """
    Carrega uma lista de nomes a partir de um TXT.
    - Um nome por linha
    - Ignora linhas vazias
    - Ignora comentarios iniciados com #
    - Normaliza (upper, remove acentos, strip, colapsa espacos)
    - Remove duplicados preservando ordem
    """
    if not os.path.exists(path):
        return []

    nomes = []
    vistos = set()
    with open(path, "r", encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#"):
                continue
            nome = _normalizar_nome(linha)
            if not nome or nome in vistos:
                continue
            vistos.add(nome)
            nomes.append(nome)
    return nomes


def _adquirir_lock(lock_path: str) -> bool:
    try:
        if os.path.exists(lock_path):
            age = datetime.now().timestamp() - os.path.getmtime(lock_path)
            if age > 30 * 60:
                os.remove(lock_path)
    except Exception:
        pass
    try:
        fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.close(fd)
        return True
    except FileExistsError:
        return False


def _liberar_lock(lock_path: str):
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass


def _carregar_txt_com_linhas(path: str):
    if not os.path.exists(path):
        return [], []
    with open(path, "r", encoding="utf-8") as f:
        linhas = f.readlines()
    nomes = []
    vistos = set()
    for linha in linhas:
        raw = linha.strip()
        if not raw or raw.startswith("#"):
            continue
        nome = _normalizar_nome(raw)
        if not nome or nome in vistos:
            continue
        vistos.add(nome)
        nomes.append(nome)
    return linhas, nomes


def _atualizar_fila_txt(path: str, nomes_processados: set[str]) -> int:
    if not os.path.exists(path):
        return 0
    with open(path, "r", encoding="utf-8") as f:
        linhas = f.readlines()

    novas_linhas = []
    removidos = 0
    for linha in linhas:
        raw = linha.strip()
        if not raw or raw.startswith("#"):
            novas_linhas.append(linha)
            continue
        nome = _normalizar_nome(raw)
        if nome in nomes_processados:
            removidos += 1
            continue
        novas_linhas.append(linha)

    nomes_restantes = [l for l in novas_linhas if l.strip() and not l.strip().startswith("#")]
    if not nomes_restantes:
        os.remove(path)
        return removidos

    with open(path, "w", encoding="utf-8") as f:
        f.writelines(novas_linhas)
    return removidos


def _escrever_documento_operacional_publico():
    os.makedirs(PUBLIC_INPUTS_DIR, exist_ok=True)
    path = os.path.join(PUBLIC_BASE_DIR, "COMO_USAR_METAX.txt")
    conteudo = (
        "OPERACAO METAXG (RAPIDO)\\n"
        f"1) Coloque o TXT em: {os.path.join(PUBLIC_INPUTS_DIR, 'cadastrar_metax.txt')}\\n"
        "2) Rode o robo normalmente (Python ou EXE).\\n"
        "3) Resolva o CAPTCHA quando solicitado.\\n"
        "4) Confira o relatorio em relatorios/ e o email (se habilitado).\\n"
        "5) O TXT se auto-limpa: nomes processados sao removidos.\\n"
        "6) Evidencias de falha de verificacao:\\n"
        "   - logs/screenshots/verify_fail_<cpf>_...png\\n"
        "   - json/verify_debug_<cpf>_...json\\n"
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(conteudo)


def _ensure_output_dirs():
    os.makedirs(os.path.join(ROOT_DIR, "logs"), exist_ok=True)
    os.makedirs(os.path.join(ROOT_DIR, "logs", "screenshots"), exist_ok=True)
    os.makedirs(os.path.join(ROOT_DIR, "relatorios"), exist_ok=True)
    os.makedirs(os.path.join(ROOT_DIR, "json"), exist_ok=True)


def _ensure_public_dirs():
    public_dirs = [
        PUBLIC_BASE_DIR,
        PUBLIC_CODE_DIR,
        PUBLIC_INPUTS_DIR,
        PUBLIC_PROCESSADOS_DIR,
        PUBLIC_ERROS_DIR,
        PUBLIC_LOGS_DIR,
        os.path.join(PUBLIC_LOGS_DIR, "screenshots") if PUBLIC_LOGS_DIR else None,
        PUBLIC_RELATORIOS_DIR,
        PUBLIC_JSON_DIR,
        PUBLIC_RELEASES_DIR,
        FOTOS_EM_PROCESSAMENTO_DIR,
        FOTOS_PROCESSADOS_DIR,
        FOTOS_ERROS_DIR,
    ]
    for path in public_dirs:
        if not path:
            continue
        try:
            os.makedirs(path, exist_ok=True)
        except Exception as e:
            logger.warn("Falha ao criar pasta publica", details={"path": path, "error": str(e)})


def _is_subpath(path: str, base: str) -> bool:
    if not path or not base:
        return False
    try:
        path_abs = os.path.abspath(path)
        base_abs = os.path.abspath(base)
        return os.path.commonpath([path_abs, base_abs]) == base_abs
    except ValueError:
        return False


def _mover_foto_para_dir(caminho_foto: str, destino_dir: str, execution_id: str) -> str | None:
    if not caminho_foto or not os.path.exists(caminho_foto):
        return None
    if not destino_dir:
        return caminho_foto
    os.makedirs(destino_dir, exist_ok=True)
    destino = os.path.join(destino_dir, os.path.basename(caminho_foto))
    if os.path.abspath(destino) == os.path.abspath(caminho_foto):
        return destino
    if os.path.exists(destino):
        base, ext = os.path.splitext(os.path.basename(caminho_foto))
        destino = os.path.join(destino_dir, f"{base}__{execution_id}{ext}")
    shutil.move(caminho_foto, destino)
    return destino


def _resolver_destino_foto(status_final: str, started_at: datetime) -> str:
    base = FOTOS_PROCESSADOS_DIR if status_final and status_final != "FAILED" else FOTOS_ERROS_DIR
    if not base:
        return ""
    data_dir = started_at.strftime("%Y-%m-%d")
    return os.path.join(base, data_dir)


def _classificar_foto_pos_processamento(
    caminho_foto: str, status_final: str, execution_id: str, started_at: datetime
):
    if not caminho_foto:
        return
    if not _is_subpath(caminho_foto, FOTOS_EM_PROCESSAMENTO_DIR):
        return
    destino = _resolver_destino_foto(status_final, started_at)
    try:
        novo_caminho = _mover_foto_para_dir(caminho_foto, destino, execution_id)
        if novo_caminho and novo_caminho != caminho_foto:
            logger.info(
                "Foto movida apos processamento",
                details={"from": caminho_foto, "to": novo_caminho, "status_final": status_final},
            )
    except Exception as e:
        logger.warn(
            "Falha ao mover foto apos processamento",
            details={"path": caminho_foto, "destino": destino, "error": str(e)},
        )


def _montar_erros_auditoria(manifest: dict) -> list[dict]:
    erros = []
    for person in manifest.get("people", []):
        if person.get("status_final") != "FAILED":
            continue
        errors = person.get("errors") or {}
        action_error = str(errors.get("action_error") or "")
        verification_error = str(errors.get("verification_error") or "")
        mensagem = action_error or verification_error
        if not mensagem:
            continue
        etapa = "Verificacao" if verification_error else "Cadastro"
        timestamps = person.get("timestamps") or {}
        timestamp = timestamps.get("saved_at") or timestamps.get("started_at") or datetime.now().isoformat()
        erros.append(
            {
                "timestamp": timestamp,
                "etapa": etapa,
                "tipo_erro": "Tecnico",
                "codigo_erro": "",
                "mensagem_resumida": mensagem[:250],
                "registro_id": person.get("cpf") or person.get("nome") or "",
                "mitigacao": "Pendente",
                "resolvido_em": "",
            }
        )
    return erros


def _parse_args():
    parser = argparse.ArgumentParser(description="RPA MetaXg")
    parser.add_argument("--txt", dest="txt_path", help="Caminho do TXT para modo manual")
    parser.add_argument("--dry-run", action="store_true", help="Nao envia e-mail (somente relatórios)")
    parser.add_argument("--no-email", action="store_true", help="Nao envia e-mail")
    parser.add_argument("--headless", action="store_true", help="Executa browser em modo headless")
    parser.add_argument("--log-level", default=os.getenv("METAX_LOG_LEVEL", "INFO"), help="INFO|DEBUG|WARN|ERROR")
    return parser.parse_args()


def main():
    args = _parse_args()
    execution_id = str(uuid.uuid4())
    started_at = datetime.now()
    output_manager = OutputManager(
        execution_id=execution_id,
        object_name=OBJECT_NAME,
        public_base_dir=PUBLIC_BASE_DIR,
        local_root=ROOT_DIR,
        started_at=started_at,
    )
    logger.configure(output_manager, execution_id, started_at, log_level=args.log_level)
    logger.set_run_status("RUNNING")
    _ensure_public_dirs()
    _ensure_output_dirs()
    _escrever_documento_operacional_publico()

    logger.info("===== INICIO PROCESSO SHAREPOINT + METAX =====")

    run_context = {
        "execution_id": execution_id,
        "object_name": OBJECT_NAME,
        "started_at": started_at.isoformat(),
        "finished_at": None,
        "duration_sec": None,
        "run_status": "RUNNING",
        "report_path": None,
        "manifest_path": None,
        "email_status": None,
        "email_error": None,
        "public_write_ok": None,
        "public_write_error": None,
        "environment": {
            "cwd": ROOT_DIR,
        },
    }

    manifest = {
        "run_context": run_context,
        "manifest_path": None,
        "totals": {
            "detected": 0,
            "people_total": 0,
            "by_outcome": {},
            "no_photo": 0,
            "unknown_outcome": 0,
        },
        "people": [],
    }

    inconsistente = False
    funcionarios = []
    try:
        txt_path = args.txt_path or os.path.join(PUBLIC_INPUTS_DIR, "cadastrar_metax.txt")
        lock_path = f"{txt_path}.lock"
        if os.path.exists(txt_path):
            lock_ok = _adquirir_lock(lock_path)
            if not lock_ok:
                logger.warn("Arquivo TXT em uso (lock ativo). Rodando em modo normal (SQL).")
                nomes_txt = []
            else:
                _, nomes_txt = _carregar_txt_com_linhas(txt_path)
        else:
            logger.info("TXT público não encontrado; rodando modo normal (SQL).")
            nomes_txt = []

        if nomes_txt:
            logger.info("Modo TXT ativo: filtrando SQL por lista manual.")
            funcionarios = buscar_funcionarios_para_cadastro(filtro_nomes=nomes_txt)
        else:
            logger.info("Modo normal: sem TXT, buscando via SQL padrao.")
            funcionarios = buscar_funcionarios_para_cadastro(filtro_nomes=None)

        if not funcionarios:
            logger.info("Nenhum funcionario encontrado para processar.")
            return

        logger.info(f"Funcionarios a processar: {len(funcionarios)}", details={"total": len(funcionarios)})

        fotos = baixar_fotos_em_lote(
            funcionarios=funcionarios,
            pasta_destino=FOTOS_EM_PROCESSAMENTO_DIR,
            pastas_busca=FOTOS_BUSCA_DIRS,
        )

        grupos = {"MECANICA": [], "ELETROMECANICA": [], "DESCONHECIDO": []}
        for func in funcionarios:
            centro_custo = func.get("CENTRO_CUSTO")
            chave = _classificar_contrato_por_centro_custo(centro_custo)
            grupos[chave].append(func)

        logger.info(
            "Resumo por contrato (centro de custo)",
            details={
                "mecanica": len(grupos["MECANICA"]),
                "eletromecanica": len(grupos["ELETROMECANICA"]),
                "desconhecido": len(grupos["DESCONHECIDO"]),
            },
        )

        nomes_processados_no_run = set()

        if grupos["DESCONHECIDO"]:
            for func in grupos["DESCONHECIDO"]:
                cpf = func["CPF"]
                cpf_limpo = "".join(filter(str.isdigit, str(cpf)))
                nome = func["NOME"]
                caminho_foto = fotos.get(cpf)
                centro_custo = func.get("CENTRO_CUSTO")
                pessoa_started_at = datetime.now().isoformat()
                registro = _criar_registro_base(nome, cpf_limpo, pessoa_started_at)
                registro["status_final"] = "FAILED"
                registro["outcome"] = OUTCOME_FAILED_ACTION
                registro["errors"]["action_error"] = f"Centro de custo desconhecido: {centro_custo}"
                logger.warn(
                    f"Centro de custo desconhecido para {nome}. Pulando cadastro.",
                    details={"cpf": cpf_limpo, "centro_custo": centro_custo},
                )
                manifest["people"].append(registro)
                _classificar_foto_pos_processamento(caminho_foto, registro["status_final"], execution_id, started_at)

        for chave in ("MECANICA", "ELETROMECANICA"):
            funcs_grupo = grupos[chave]
            if not funcs_grupo:
                continue

            contrato_value, contrato_label = _resolver_contrato_config(chave)
            if not (contrato_value or contrato_label):
                raise ValueError(
                    f"Contrato {chave} nao configurado. "
                    f"Defina METAX_CONTRATO_{chave}_VALUE ou METAX_CONTRATO_{chave}_LABEL no .env"
                )

            logger.info(f"Iniciando sessao para contrato {chave}...")
            p, browser, page = iniciar_sessao(
                headless=args.headless,
                contrato_value=contrato_value,
                contrato_label=contrato_label,
            )

            try:
                rascunhos_existentes = obter_todos_rascunhos(page)

                for func in funcs_grupo:
                    cpf = func["CPF"]
                    cpf_limpo = "".join(filter(str.isdigit, str(cpf)))
                    nome = func["NOME"]

                    pessoa_started_at = datetime.now().isoformat()
                    registro = _criar_registro_base(nome, cpf_limpo, pessoa_started_at)
                    caminho_foto = fotos.get(cpf)

                    if cpf_limpo in rascunhos_existentes:
                        logger.info(f"Funcionario {nome} ja consta nos rascunhos (CACHE). Pulando...", details={"cpf": cpf})
                        registro["attempted"] = False
                        registro["status_final"] = "SKIPPED"
                        registro["outcome"] = OUTCOME_SKIPPED_ALREADY_EXISTS
                        registro["errors"]["action_error"] = "Ignorado: rascunho ja existente (cache)."
                        nomes_processados_no_run.add(_normalizar_nome(nome))
                        manifest["people"].append(registro)
                        _classificar_foto_pos_processamento(caminho_foto, registro["status_final"], execution_id, started_at)
                        continue

                    if caminho_foto:
                        logger.info(f"Foto pronta para {nome}", details={"cpf": cpf, "foto": caminho_foto})
                    else:
                        logger.warn(f"Foto nao encontrada para {nome}", details={"cpf": cpf})
                        registro["no_photo"] = True

                    logger.info(f"Iniciando cadastro de {nome} ({cpf})", details={"funcionario": nome, "cpf": cpf})

                    try:
                        action = cadastrar_funcionario(page, func, output_manager, caminho_foto)
                    except Exception as e:
                        logger.error(f"Falha ao cadastrar {nome}: {e}", details={"cpf": cpf, "erro": str(e)})
                        registro["attempted"] = False
                        registro["action_saved"] = False
                        registro["status_final"] = "FAILED"
                        registro["outcome"] = OUTCOME_FAILED_ACTION
                        registro["errors"]["action_error"] = str(e)

                        try:
                            page.goto("https://portal.metax.ind.br/", timeout=5000)
                        except Exception:
                            pass

                        manifest["people"].append(registro)
                        _classificar_foto_pos_processamento(caminho_foto, registro["status_final"], execution_id, started_at)
                        continue

                    # Blindagem do contrato de retorno do action
                    if not isinstance(action, dict):
                        action = {"attempted": False, "saved": False, "no_photo": False, "error": "Retorno invalido", "detail": ""}
                    action = {
                        "attempted": bool(action.get("attempted", False)),
                        "saved": bool(action.get("saved", False)),
                        "no_photo": bool(action.get("no_photo", False)),
                        "error": str(action.get("error", "")),
                        "detail": str(action.get("detail", "")),
                    }

                    registro["attempted"] = action["attempted"]
                    registro["action_saved"] = action["saved"]
                    if action.get("no_photo"):
                        registro["no_photo"] = True
                    if registro["attempted"]:
                        nomes_processados_no_run.add(_normalizar_nome(nome))

                    if registro["action_saved"]:
                        registro["timestamps"]["saved_at"] = datetime.now().isoformat()
                        logger.info(f"[VERIFY] start cpf={cpf_limpo}, nome={nome}")
                        try:
                            verificado, detalhe = verificar_cadastro(page, func, output_manager)
                        except Exception as e:
                            verificado, detalhe = False, f"Erro na verificacao: {e}"
                        logger.info(f"[VERIFY] result cpf={cpf_limpo} verified={bool(verificado)} detail={detalhe}")

                        registro["verified"] = bool(verificado)
                        if verificado:
                            registro["timestamps"]["verified_at"] = datetime.now().isoformat()
                            registro["status_final"] = "SUCCESS"
                            registro["outcome"] = OUTCOME_VERIFIED_SUCCESS
                            rascunhos_existentes.add(cpf_limpo)
                            logger.info("Cache de rascunhos atualizado.", details={"cpf": cpf_limpo})
                        else:
                            logger.warn(f"Verificacao falhou para {nome}: {detalhe}", details={"cpf": cpf, "motivo": detalhe})
                            registro["status_final"] = "FAILED"
                            if detalhe and detalhe.lower().startswith("erro na verificacao"):
                                registro["outcome"] = OUTCOME_FAILED_VERIFICATION
                            else:
                                registro["outcome"] = OUTCOME_SAVED_NOT_VERIFIED
                            registro["errors"]["verification_error"] = detalhe or "CPF nao encontrado na lista de rascunhos."
                            inconsistente = True
                    else:
                        registro["status_final"] = "FAILED"
                        registro["outcome"] = OUTCOME_FAILED_ACTION
                        registro["errors"]["action_error"] = action.get("error") or "Falha ao salvar rascunho."

                    manifest["people"].append(registro)
                    _classificar_foto_pos_processamento(caminho_foto, registro["status_final"], execution_id, started_at)
            finally:
                if browser:
                    logger.info("Fechando navegador...")
                    browser.close()
                if p:
                    p.stop()

    finally:
        try:
            if "lock_ok" in locals() and lock_ok and os.path.exists(txt_path):
                removidos = _atualizar_fila_txt(txt_path, nomes_processados_no_run)
                logger.info(f"TXT fila atualizado. Nomes removidos: {removidos}")
        finally:
            if "lock_ok" in locals() and lock_ok:
                _liberar_lock(lock_path)

        finished_at = datetime.now()
        run_context["finished_at"] = finished_at.isoformat()
        run_context["duration_sec"] = int((finished_at - started_at).total_seconds())

        totals = compute_totals(manifest["people"], detected=len(funcionarios))
        manifest["totals"] = totals

        if (
            inconsistente
            or totals["by_outcome"].get(OUTCOME_SAVED_NOT_VERIFIED, 0) > 0
            or totals["by_outcome"].get(OUTCOME_FAILED_ACTION, 0) > 0
            or totals["by_outcome"].get(OUTCOME_FAILED_VERIFICATION, 0) > 0
        ):
            run_context["run_status"] = "INCONSISTENT"
        else:
            run_context["run_status"] = "CONSISTENT"

        run_context["public_write_ok"] = output_manager.public_write_ok
        run_context["public_write_error"] = output_manager.public_write_error
        logger.set_run_status(run_context["run_status"])

        data_str = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        manifest_filename = f"manifest_{data_str}__{execution_id}.json"
        final_manifest_path = output_manager.get_local_path(KIND_JSON, manifest_filename)
        run_context["manifest_path"] = final_manifest_path
        manifest["manifest_path"] = final_manifest_path

        logger.info("Gerando relatorios...")
        report_path = gerar_relatorio_txt(manifest, output_manager)
        run_context["report_path"] = report_path
        resumo_path = gerar_resumo_execucao_md(manifest, output_manager)
        run_context["resumo_path"] = resumo_path
        relatorio_json_path = gerar_relatorio_json(manifest, output_manager)
        run_context["relatorio_json_path"] = relatorio_json_path
        diagnostico_path = gerar_diagnostico_ultima_execucao(manifest, output_manager)
        run_context["diagnostico_path"] = diagnostico_path

        manifest_partial_filename = f"manifest_partial_{data_str}__{execution_id}.json"
        manifest_partial_path = output_manager.write_json(KIND_JSON, manifest_partial_filename, manifest)

        if not args.dry_run and not args.no_email:
            try:
                email_status = enviar_relatorio_email(
                    manifest,
                    report_path=report_path,
                    manifest_path=final_manifest_path,
                    partial_manifest_path=manifest_partial_path,
                    attachment_paths=[report_path, manifest_partial_path],
                )
                run_context["email_status"] = email_status
            except Exception as e:
                run_context["email_status"] = "FAILED"
                run_context["email_error"] = str(e)
        else:
            if args.dry_run:
                run_context["email_status"] = OUTCOME_SKIPPED_DRY_RUN
            else:
                run_context["email_status"] = OUTCOME_SKIPPED_EMAIL_DISABLED

        output_manager.write_json(KIND_JSON, manifest_filename, manifest)
        try:
            total_sucesso = sum(1 for p in manifest.get("people", []) if p.get("status_final") == "SUCCESS")
            total_erro = sum(1 for p in manifest.get("people", []) if p.get("status_final") == "FAILED")
            total_processado = total_sucesso + total_erro
            skipped = len(manifest.get("people", [])) - total_processado
            observacoes = ""
            if skipped > 0:
                observacoes = f"Skipped={skipped}"

            audit_run_data = {
                "run_id": execution_id,
                "started_at": started_at,
                "finished_at": finished_at,
                "duration_sec": run_context.get("duration_sec"),
                "total_processado": total_processado,
                "total_sucesso": total_sucesso,
                "total_erro": total_erro,
                "erros_auto_mitigados": 0,
                "erros_manuais": 0,
                "ambiente": os.getenv("METAX_ENV", "PROD"),
                "observacoes": observacoes,
                "commit_hash": os.getenv("GIT_COMMIT") or os.getenv("GITHUB_SHA") or "",
                "build_id": os.getenv("BUILD_ID") or os.getenv("GITHUB_RUN_ID") or "",
            }
            audit_errors = _montar_erros_auditoria(manifest)
            audit_result = log_run(audit_run_data, audit_errors)
            logger.info("Auditoria Excel atualizada", details=audit_result)
        except Exception as e:
            logger.error("Falha ao atualizar auditoria Excel", details={"error": str(e)})
        logger.flush()
        logger.info("===== FIM PROCESSO =====")


if __name__ == "__main__":
    main()
