from rpa_metax import iniciar_sessao, cadastrar_funcionario
from sharepoint import baixar_fotos_em_lote
import pyodbc
from datetime import datetime
from custom_logger import logger

from config import (
    DB_DRIVER, DB_SERVER, DB_NAME, DB_USER, DB_PASSWORD, PASTA_FOTOS
)

def obter_conexao():
    return pyodbc.connect(
        f"DRIVER={{{DB_DRIVER}}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASSWORD}"
    )

def buscar_funcionarios_para_cadastro(data_admissao=None):
    if not data_admissao:
        data_admissao = datetime.now().strftime("%Y-%m-%d")

    logger.info(f"Buscando funcionários com admissão em: {data_admissao}", details={"data_admissao": data_admissao})

    sql = f"""
        DECLARE @DataAdmissao DATE = '{data_admissao}';

        SELECT 
            P.NOME,
            P.CPF,
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
            CAST(F.DATAADMISSAO AS DATE) = @DataAdmissao
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

def main():
    logger.info("===== INÍCIO PROCESSO SHAREPOINT + METAX =====")

    funcionarios = buscar_funcionarios_para_cadastro()

    if not funcionarios:
        logger.info("Nenhum funcionário para cadastrar hoje.")
        return

    logger.info(f"Funcionários a processar: {len(funcionarios)}", details={"total": len(funcionarios)})

    fotos = baixar_fotos_em_lote(
        funcionarios=funcionarios,
        pasta_destino=PASTA_FOTOS
    )

    p, browser, page = iniciar_sessao()

    try:
        for func in funcionarios:
            cpf = func["CPF"]
            caminho_foto = fotos.get(cpf)

            if caminho_foto:
                logger.info(f"Foto pronta para {func['NOME']}", details={"cpf": cpf, "foto": caminho_foto})
            else:
                logger.warn(f"Foto não encontrada para {func['NOME']}", details={"cpf": cpf})

            logger.info(f"Iniciando cadastro de {func['NOME']} ({cpf})", details={"funcionario": func['NOME'], "cpf": cpf})
            
            try:
                cadastrar_funcionario(page, func, caminho_foto)
            except Exception as e:
                logger.error(f"Falha ao cadastrar {func['NOME']}: {e}", details={"cpf": cpf, "erro": str(e)})
                # Tenta recuperar navegação
                try:
                   # Tenta ir para a home se der erro
                   page.goto("https://portal.metax.ind.br/", timeout=5000)
                except:
                    pass

    finally:
        logger.info("Fechando navegador...")
        browser.close()
        p.stop()


    logger.info("===== FIM PROCESSO =====")


if __name__ == "__main__":
    main()