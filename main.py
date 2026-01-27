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

    if filtro_nomes and len(filtro_nomes) > 0:
        logger.info(
            f"MODO FILTRO ATIVADO: Buscando {len(filtro_nomes)} funcionario(s) especifico(s). Datas serao ignoradas.",
            details={"nomes": filtro_nomes},
        )
        lista_sql = ", ".join([f"'{nome.strip().upper()}'" for nome in filtro_nomes])
        where_clause = f"P.NOME IN ({lista_sql})"
        datas_vars = f"""
            DECLARE @DataFim DATE = '{data_admissao}';
            DECLARE @DataInicio DATE = DATEADD(DAY, -{DIAS_RETROATIVOS}, @DataFim);
        """
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

        datas_vars = f"""
            DECLARE @DataFim DATE = '{data_admissao}';
            DECLARE @DataInicio DATE = DATEADD(DAY, -{DIAS_RETROATIVOS}, @DataFim);
        """
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
        cursor.execute(sql)
        colunas = [c[0] for c in cursor.description]
        return [dict(zip(colunas, row)) for row in cursor.fetchall()]


from notification import enviar_relatorio_email
from reporting import gerar_relatorio_txt


def main():
    execution_id = str(uuid.uuid4())
    started_at = datetime.now()
    output_manager = OutputManager(
        execution_id=execution_id,
        object_name=OBJECT_NAME,
        public_base_dir=PUBLIC_BASE_DIR,
        local_root=ROOT_DIR,
        started_at=started_at,
    )
    logger.configure(output_manager, execution_id, started_at)

    logger.info("===== INICIO PROCESSO SHAREPOINT + METAX =====")

    manifest = {
        "execution_id": execution_id,
        "object_name": OBJECT_NAME,
        "started_at": started_at.isoformat(),
        "finished_at": None,
        "run_status": "OK",
        "public_write_ok": None,
        "public_write_error": None,
        "totals": {
            "detected": 0,
            "processed_success": 0,
            "ignored": 0,
            "failed": 0,
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
        funcionarios_pontuais = [
            "ALEXSANDRO DE ALMEIDA CARDOSO",
            "ALEX DA SILVA ALEIXO",
            "PAULO GARCIAS PIMENTA LOPES",
            "RONIERE ASSUNCAO",
            "JOSE JUNIOR DA SILVA LIMA",
            "TAYRON HENRIKE DOS SANTOS AMORIM",
            "WANDECARLOS DE ASSUNCAO OLIVEIRA",
        ]
        # funcionarios_pontuais = None

        funcionarios = buscar_funcionarios_para_cadastro(filtro_nomes=funcionarios_pontuais)

        if not funcionarios:
            logger.info("Nenhum funcionario encontrado para processar.")
            return

        logger.info(f"Funcionarios a processar: {len(funcionarios)}", details={"total": len(funcionarios)})

        fotos = baixar_fotos_em_lote(
            funcionarios=funcionarios,
            pasta_destino=PASTA_FOTOS,
        )

        p, browser, page = iniciar_sessao()

        rascunhos_existentes = obter_todos_rascunhos(page)

        for func in funcionarios:
            cpf = func["CPF"]
            cpf_limpo = "".join(filter(str.isdigit, str(cpf)))
            nome = func["NOME"]

            registro = {
                "nome": nome,
                "cpf": cpf_limpo,
                "status": "FAILED",
                "verified": False,
                "verification_detail": "",
                "action_result": {},
                "error": "",
                "no_photo": False,
            }

            if cpf_limpo in rascunhos_existentes:
                logger.info(f"Funcionario {nome} ja consta nos rascunhos (CACHE). Pulando...", details={"cpf": cpf})
                registro["status"] = "IGNORED"
                registro["verification_detail"] = "Ignorado: rascunho ja existente (cache)."
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
                registro["status"] = "FAILED"
                registro["error"] = str(e)
                registro["action_result"] = {"attempted": False, "saved": False, "error": str(e)}

                try:
                    page.goto("https://portal.metax.ind.br/", timeout=5000)
                except Exception:
                    pass

                manifest["people"].append(registro)
                continue

            registro["action_result"] = action
            if action.get("no_photo"):
                registro["no_photo"] = True

            if action.get("saved"):
                try:
                    verificado, detalhe = verificar_cadastro(page, func)
                except Exception as e:
                    verificado = False
                    detalhe = f"Erro na verificacao: {e}"

                registro["verified"] = bool(verificado)
                registro["verification_detail"] = detalhe or ""

                if verificado:
                    registro["status"] = "SUCCESS"
                    rascunhos_existentes.add(cpf_limpo)
                    logger.info("Cache de rascunhos atualizado.", details={"cpf": cpf_limpo})
                else:
                    registro["status"] = "FAILED"
                    registro["error"] = "Cadastro nao verificado."
                    inconsistente = True
            else:
                registro["status"] = "FAILED"
                registro["verification_detail"] = "Salvamento nao confirmado."
                registro["error"] = action.get("error") or "Falha ao salvar rascunho."

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
            "processed_success": len([p for p in manifest["people"] if p["status"] == "SUCCESS" and p["verified"]]),
            "ignored": len([p for p in manifest["people"] if p["status"] == "IGNORED"]),
            "failed": len([p for p in manifest["people"] if p["status"] == "FAILED"]),
            "no_photo": len([p for p in manifest["people"] if p.get("no_photo")]),
        }
        manifest["totals"] = totals

        if inconsistente:
            manifest["run_status"] = "INCONSISTENT"

        manifest["public_write_ok"] = output_manager.public_write_ok
        manifest["public_write_error"] = output_manager.public_write_error

        data_str = started_at.strftime("%Y-%m-%d_%H-%M-%S")
        manifest_filename = f"manifest_{data_str}__{execution_id}.json"
        output_manager.write_json(KIND_JSON, manifest_filename, manifest)

        logger.info("Gerando relatorios...")
        gerar_relatorio_txt(manifest, output_manager)
        enviar_relatorio_email(manifest)
        logger.info("===== FIM PROCESSO =====")


if __name__ == "__main__":
    main()
