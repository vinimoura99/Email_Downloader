import win32com.client
import os
import sys
import time
from datetime import datetime

# =============================================================================
# CONFIGURAÇÕES DE BUSCA (ALTERE AQUI PARA MUDAR O ALVO)
# =============================================================================

# 1. Termo de busca (O que procurar no remetente ou no assunto)
TERMO_BUSCA = "nome do arquivo fornecido em massa " 

# 2. Extensão do arquivo que deseja baixar
EXTENSAO_ALVO = ".pdf"

# 3. Nome da pasta principal onde os arquivos serão salvos
NOME_PASTA_RAIZ = "Downloads_Outlook"

# 4. Quantidade de e-mails recentes para verificar
LIMITE_MENSAGENS = 150
# =============================================================================

def obter_caminho_base():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def executar_busca_outlook():
    try:
        base_path = obter_camin_base()
        hoje_str = datetime.now().strftime("%d/%m/%Y")
        hoje_pasta = datetime.now().strftime("%d-%m-%Y")
        
        # Define o caminho final: Downloads_Outlook / 18-03-2026
        caminho_final = os.path.join(base_path, NOME_PASTA_RAIZ, hoje_pasta)

        if not os.path.exists(caminho_final):
            os.makedirs(caminho_final)

        print(f"🔗 Conectando ao serviço de e-mail...")
        
        # Conecta ao Outlook (Tenta usar instância aberta primeiro)
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except:
            outlook = win32com.client.Dispatch("Outlook.Application")
            
        ns = outlook.GetNamespace("MAPI")
        inbox = ns.GetDefaultFolder(6) # Pasta 6 = Inbox (Caixa de Entrada)
        
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True) # Mais recentes primeiro

        print(f"📅 Verificando mensagens de hoje: {hoje_str}")
        print(f"🔍 Termo de busca: '{TERMO_BUSCA}'")

        count_arquivos = 0
        
        for i in range(1, min(len(messages), LIMITE_MENSAGENS)):
            try:
                msg = messages.Item(i)
                
                # Filtra pela data de hoje
                if msg.ReceivedTime.strftime("%d/%m/%Y") != hoje_str:
                    continue 

                remetente = str(msg.SenderName).lower()
                assunto = str(msg.Subject).lower()

                # Verifica se o termo de busca está no remetente ou assunto
                if TERMO_BUSCA.lower() in remetente or TERMO_BUSCA.lower() in assunto:
                    if msg.Attachments.Count > 0:
                        for j in range(1, msg.Attachments.Count + 1):
                            att = msg.Attachments.Item(j)
                            
                            if att.FileName.lower().endswith(EXTENSAO_ALVO.lower()):
                                nome_arq = att.FileName
                                destino = os.path.join(caminho_final, nome_arq)
                                
                                if not os.path.exists(destino):
                                    att.SaveAsFile(destino)
                                    print(f"✅ Baixado: {nome_arq}")
                                    count_arquivos += 1
                                else:
                                    print(f"ℹ️ Já ignorado (existente): {nome_arq}")
            except:
                continue

        print(f"\n✨ Processo finalizado!")
        print(f"📊 Total de arquivos extraídos: {count_arquivos}")
        print(f"📁 Local: {caminho_final}")

    except Exception as e:
        print(f"❌ Erro Crítico: {e}")

if __name__ == "__main__":
    executar_busca_outlook()
    print("-" * 30)
    input("Pressione ENTER para fechar...")