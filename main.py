from rpa_metax import iniciar_sessao, cadastrar_funcionario, obter_todos_rascunhos
from sharepoint import baixar_fotos_em_lote
import pyodbc
from datetime import datetime
from custom_logger import logger

from config import (
    DB_DRIVER, DB_SERVER, DB_NAME, DB_USER, DB_PASSWORD, PASTA_FOTOS, DIAS_RETROATIVOS
)

def obter_conexao() -> pyodbc.Connection:
    """Estabelece conexão com o banco de dados SQL Server."""
    return pyodbc.connect(
        f"DRIVER={{{DB_DRIVER}}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASSWORD}"
    )

def buscar_funcionarios_para_cadastro(data_admissao: str = None) -> list[dict]:
    """
    Busca funcionários admitidos na data especificada (ou hoje, se data não fornecida).
    Se DIAS_RETROATIVOS > 0, busca no intervalo [Data - Dias, Data].
    
    Args:
        data_admissao (str, optional): Data final de admissão no formato YYYY-MM-DD. Defaults to None (hoje).
        
    Returns:
        list[dict]: Lista de dicionários com os dados dos funcionários.
    """
    if not data_admissao:
        data_admissao = datetime.now().strftime("%Y-%m-%d")

    # Mensagem de log ajustada
    if DIAS_RETROATIVOS > 0:
        logger.info(f"Buscando funcionários com admissão entre {data_admissao} e {DIAS_RETROATIVOS} dia(s) antes.", details={"data_ref": data_admissao, "retroativos": DIAS_RETROATIVOS})
    else:
        logger.info(f"Buscando funcionários com admissão em: {data_admissao}", details={"data_admissao": data_admissao})

    sql = f"""
        DECLARE @DataFim DATE = '{data_admissao}';
        DECLARE @DataInicio DATE = DATEADD(DAY, -{DIAS_RETROATIVOS}, @DataFim);

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
            CAST(F.DATAADMISSAO AS DATE) BETWEEN @DataInicio AND @DataFim
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
    logger.info("===== INÍCIO PROCESSO SHAREPOINT + METAX =====")

    funcionarios = buscar_funcionarios_para_cadastro()

    if not funcionarios:
        logger.info("Nenhum funcionário para cadastrar hoje.")
        return

    logger.info(f"Funcionários a processar: {len(funcionarios)}", details={"total": len(funcionarios)})

    # Estatísticas
    stats = {
        "total": len(funcionarios),
        "sucesso": [],
        "falha": [],
        "sem_foto": [],
        "duplicados": []
    }

    fotos = baixar_fotos_em_lote(
        funcionarios=funcionarios,
        pasta_destino=PASTA_FOTOS
    )

    p, browser, page = iniciar_sessao()

    try:
        # 1. Obter lista de rascunhos JÁ EXISTENTES de uma vez só
        rascunhos_existentes = obter_todos_rascunhos(page)
        
        for func in funcionarios:
            cpf = func["CPF"]
            cpf_limpo = ''.join(filter(str.isdigit, str(cpf)))
            nome = func["NOME"]
            
            # 2. Verifica se JÁ ESTÁ na lista de rascunhos (Cache Local)
            if cpf_limpo in rascunhos_existentes:
                logger.info(f"Funcionário {nome} já consta nos rascunhos (CACHE). Pulando...", details={"cpf": cpf})
                stats["duplicados"].append(nome)
                continue

            # Se não está cadastrado, segue o fluxo
            caminho_foto = fotos.get(cpf)

            if caminho_foto:
                logger.info(f"Foto pronta para {nome}", details={"cpf": cpf, "foto": caminho_foto})
            else:
                logger.warn(f"Foto não encontrada para {nome}", details={"cpf": cpf})
                stats["sem_foto"].append(nome) # Adiciona à lista de sem foto

            logger.info(f"Iniciando cadastro de {nome} ({cpf})", details={"funcionario": nome, "cpf": cpf})
            
            try:
                # Agora retorna STATUS (str)
                # IMPORTANTE: navegar_para_cadastro não verifica mais duplicidade interna
                status = cadastrar_funcionario(page, func, caminho_foto)
                
                if status == 'SUCESSO':
                    stats["sucesso"].append(nome)
                elif status == 'DUPLICADO': 
                    # Fallback caso a verificação em cache falhe e a função interna detecte (se mantivermos lógica lá)
                    # Mas como removemos a varredura interna, isso só aconteceria se o cadastro falhasse por duplicidade no save
                    stats["duplicados"].append(nome)
                else: # ERRO
                    stats["falha"].append({"nome": nome, "motivo": "Erro no processo de cadastro"})


            except Exception as e:
                logger.error(f"Falha ao cadastrar {nome}: {e}", details={"cpf": cpf, "erro": str(e)})
                # Salva o objeto COMPLETO para retry
                func_copia = func.copy()
                func_copia["motivo_erro"] = str(e)
                stats["falha"].append(func_copia)
                
                # Tenta recuperar navegação
                try:
                   page.goto("https://portal.metax.ind.br/", timeout=5000)
                except:
                    pass

        # === TENTATIVA DE RETRY (REPESCAGEM) ===
        if stats["falha"]:
            logger.info(f"⚠️ Iniciando REPESCAGEM de {len(stats['falha'])} falhas...")
            falhas_para_retry = stats["falha"][:] # Copia a lista
            stats["falha"] = [] # Limpa para preencher com o resultado do retry
            
            for func_retry in falhas_para_retry:
                nome = func_retry["NOME"]
                cpf = func_retry["CPF"]
                logger.info(f"🔄 Retentando cadastro de {nome}...")
                
                try:
                    caminho_foto = fotos.get(cpf)
                    status = cadastrar_funcionario(page, func_retry, caminho_foto)
                    
                    if status == 'SUCESSO':
                        logger.info(f"✅ Recuperado com sucesso: {nome}")
                        stats["sucesso"].append(nome)
                    elif status == 'DUPLICADO':
                         stats["duplicados"].append(nome)
                    else:
                        stats["falha"].append({"nome": nome, "motivo": "Falha no Retry"})
                except Exception as e_retry:
                     logger.error(f"❌ Falha definitiva em {nome}: {e_retry}")
                     stats["falha"].append({"nome": nome, "motivo": f"Falha no Retry: {str(e_retry)}"})
                     try:
                        page.goto("https://portal.metax.ind.br/", timeout=5000)
                     except:
                        pass
            
    finally:
        logger.info("Fechando navegador...")
        browser.close()
        p.stop()
        
        logger.info("Gerando relatórios...")
        gerar_relatorio_txt(stats)
        enviar_relatorio_email(stats)
        
        logger.info("===== FIM PROCESSO =====")


    logger.info("===== FIM PROCESSO =====")


if __name__ == "__main__":
    main()