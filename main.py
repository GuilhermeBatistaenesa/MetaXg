import argparse
import os
import unicodedata
import uuid
from datetime import datetime

import pyodbc
from custom_logger import logger
from output_manager import OutputManager, KIND_JSON
from rpa_metax import iniciar_sessao, cadastrar_funcionario, obter_todos_rascunhos, verificar_cadastro
from sharepoint import baixar_fotos_em_lote

from config import (
    DB_DRIVER, DB_SERVER, DB_NAME, DB_USER, DB_PASSWORD, PASTA_FOTOS, DIAS_RETROATIVOS,
    ROOT_DIR, PUBLIC_BASE_DIR, OBJECT_NAME
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
from reporting import gerar_relatorio_txt


def _normalizar_nome(raw: str) -> str:
    if not raw:
        return ""
    texto = raw.strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join([c for c in texto if not unicodedata.combining(c)])
    texto = " ".join(texto.split())
    return texto.upper()


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


def _ensure_output_dirs():
    os.makedirs(os.path.join(ROOT_DIR, "logs"), exist_ok=True)
    os.makedirs(os.path.join(ROOT_DIR, "logs", "screenshots"), exist_ok=True)
    os.makedirs(os.path.join(ROOT_DIR, "relatorios"), exist_ok=True)
    os.makedirs(os.path.join(ROOT_DIR, "json"), exist_ok=True)


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
    _ensure_output_dirs()

    logger.info("===== INICIO PROCESSO SHAREPOINT + METAX =====")

    manifest = {
        "execution_id": execution_id,
        "object_name": OBJECT_NAME,
        "started_at": started_at.isoformat(),
        "finished_at": None,
        "run_status": "CONSISTENT",
        "public_write_ok": None,
        "public_write_error": None,
        "totals": {
            "detected": 0,
            "action_saved": 0,
            "verified_success": 0,
            "saved_not_verified": 0,
            "failed_action": 0,
            "failed_verification": 0,
            "no_photo": 0,
        },
        "people": [],
    }

    inconsistente = False
    funcionarios = []
    p = None
    browser = None
    page = None

    try:
        txt_path = args.txt_path or os.path.join(ROOT_DIR, "inputs", "cadastrar_metax.txt")
        nomes_txt = carregar_lista_nomes_txt(txt_path)

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
            pasta_destino=PASTA_FOTOS,
        )

        p, browser, page = iniciar_sessao(headless=args.headless)

        rascunhos_existentes = obter_todos_rascunhos(page)

        for func in funcionarios:
            cpf = func["CPF"]
            cpf_limpo = "".join(filter(str.isdigit, str(cpf)))
            nome = func["NOME"]

            pessoa_started_at = datetime.now().isoformat()
            registro = {
                "nome": nome,
                "cpf": cpf_limpo,
                "attempted": False,
                "action_saved": False,
                "verified": False,
                "status_final": "FAILED",
                "outcome": "FAILED_ACTION",
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

            if cpf_limpo in rascunhos_existentes:
                logger.info(f"Funcionario {nome} ja consta nos rascunhos (CACHE). Pulando...", details={"cpf": cpf})
                registro["attempted"] = False
                registro["status_final"] = "FAILED"
                registro["outcome"] = "FAILED_ACTION"
                registro["errors"]["action_error"] = "Ignorado: rascunho ja existente (cache)."
                manifest["people"].append(registro)
                continue

            caminho_foto = fotos.get(cpf)
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
                registro["outcome"] = "FAILED_ACTION"
                registro["errors"]["action_error"] = str(e)

                try:
                    page.goto("https://portal.metax.ind.br/", timeout=5000)
                except Exception:
                    pass

                manifest["people"].append(registro)
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
                    registro["outcome"] = "VERIFIED_SUCCESS"
                    rascunhos_existentes.add(cpf_limpo)
                    logger.info("Cache de rascunhos atualizado.", details={"cpf": cpf_limpo})
                else:
                    logger.warn(f"Verificacao falhou para {nome}: {detalhe}", details={"cpf": cpf, "motivo": detalhe})
                    registro["status_final"] = "FAILED"
                    if detalhe and detalhe.lower().startswith("erro na verificacao"):
                        registro["outcome"] = "FAILED_VERIFICATION"
                    else:
                        registro["outcome"] = "SAVED_NOT_VERIFIED"
                    registro["errors"]["verification_error"] = detalhe or "CPF nao encontrado na lista de rascunhos."
                    inconsistente = True
            else:
                registro["status_final"] = "FAILED"
                registro["outcome"] = "FAILED_ACTION"
                registro["errors"]["action_error"] = action.get("error") or "Falha ao salvar rascunho."

            manifest["people"].append(registro)

    finally:
        if browser:
            logger.info("Fechando navegador...")
            browser.close()
        if p:
            p.stop()

        finished_at = datetime.now()
        manifest["finished_at"] = finished_at.isoformat()

        totals = {
            "detected": len(funcionarios),
            "action_saved": len([p for p in manifest["people"] if p.get("action_saved")]),
            "verified_success": len([p for p in manifest["people"] if p.get("outcome") == "VERIFIED_SUCCESS"]),
            "saved_not_verified": len([p for p in manifest["people"] if p.get("outcome") == "SAVED_NOT_VERIFIED"]),
            "failed_action": len([p for p in manifest["people"] if p.get("outcome") == "FAILED_ACTION"]),
            "failed_verification": len([p for p in manifest["people"] if p.get("outcome") == "FAILED_VERIFICATION"]),
            "no_photo": len([p for p in manifest["people"] if p.get("no_photo")]),
        }
        manifest["totals"] = totals

        if inconsistente or totals["saved_not_verified"] > 0:
            manifest["run_status"] = "INCONSISTENT"

        manifest["public_write_ok"] = output_manager.public_write_ok
        manifest["public_write_error"] = output_manager.public_write_error

        data_str = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        manifest_filename = f"manifest_{data_str}__{execution_id}.json"
        output_manager.write_json(KIND_JSON, manifest_filename, manifest)

        logger.info("Gerando relatorios...")
        gerar_relatorio_txt(manifest, output_manager)
        if not args.dry_run and not args.no_email:
            enviar_relatorio_email(manifest)
        logger.info("===== FIM PROCESSO =====")


if __name__ == "__main__":
    main()
