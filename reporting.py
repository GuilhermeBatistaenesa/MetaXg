import os
from datetime import datetime
from config import RELATORIOS_DIR
from custom_logger import logger

def gerar_relatorio_txt(stats: dict):
    """
    Gera um relatório TXT detalhado da execução na pasta relatorios.
    """
    try:
        data_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        nome_arquivo = f"relatorio_execucao_{data_str}.txt"
        caminho_completo = os.path.join(RELATORIOS_DIR, nome_arquivo)
        
        with open(caminho_completo, "w", encoding="utf-8") as f:
            f.write(f"RELATÓRIO DE EXECUÇÃO RPA METAX\n")
            f.write(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write("="*50 + "\n\n")
            
            f.write(f"RESUMO GERAL:\n")
            f.write(f"- Total Detectado no RM: {stats['total']}\n")
            f.write(f"- Já Cadastrados (Ignorados): {len(stats['duplicados'])}\n")
            f.write(f"- Processados com Sucesso: {len(stats['sucesso'])}\n")
            f.write(f"- Erros / Falhas: {len(stats['falha'])}\n")
            f.write(f"- Sem Foto (Processados): {len(stats['sem_foto'])}\n")
            f.write("\n" + "="*50 + "\n\n")
            
            if stats['duplicados']:
                f.write("LISTA DE JÁ CADASTRADOS (IGNORADOS):\n")
                for nome in stats['duplicados']:
                    f.write(f" - {nome}\n")
                f.write("\n")

            if stats['sucesso']:
                f.write("LISTA DE SUCESSO:\n")
                for nome in stats['sucesso']:
                    f.write(f" - {nome}\n")
                f.write("\n")

            if stats['sem_foto']:
                f.write("LISTA SEM FOTO (MAS CADASTRADOS):\n")
                for nome in stats['sem_foto']:
                    f.write(f" - {nome}\n")
                f.write("\n")
                
            if stats['falha']:
                f.write("LISTA DE ERROS:\n")
                for item in stats['falha']:
                    nome = item.get('NOME') or item.get('nome') or "Desconhecido"
                    motivo = item.get('motivo_erro') or item.get('motivo') or "Motivo não especificado"
                    f.write(f" - {nome}: {motivo}\n")
                f.write("\n")
                
        logger.info(f"Relatório TXT gerado em: {caminho_completo}")
        return caminho_completo
        
    except Exception as e:
        logger.error(f"Erro ao gerar relatório TXT: {e}")
        return None
