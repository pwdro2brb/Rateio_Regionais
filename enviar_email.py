import os
import win32com.client as win32
from datetime import datetime, timedelta

def criar_rascunhos_correios():
    # 1. Definir o caminho base
    caminho_base = r"\\Bhz-fls-app1\mrvbh\Gerência Administrativa\Pública\NUCLEO DE CONTRATOS E APOIO A GESTÃO\CONTRATOS\Contratos Serviços\1. CORREIOS\2. Faturamento\2026"
    
    # 2. Encontrar a pasta do mês mais recente
    # Lista apenas as pastas dentro do diretório base
    pastas_meses = [f for f in os.listdir(caminho_base) if os.path.isdir(os.path.join(caminho_base, f))]
    
    if not pastas_meses:
        print("Nenhuma pasta de mês encontrada no diretório.")
        return

    # Ordena as pastas (como elas começam com números ex: "07 -Julho", a ordem alfabética funciona perfeitamente)
    pastas_meses.sort()
    pasta_mes_recente = pastas_meses[-1] # Pega a última da lista
    caminho_mes_recente = os.path.join(caminho_base, pasta_mes_recente)
    
    # Extrai apenas o nome do mês para o assunto (Ex: "07 -Julho" vira "Julho")
    nome_mes = pasta_mes_recente.split("-")[-1].strip()

    print(f"Pasta mais recente encontrada: {pasta_mes_recente}")

    # Dicionário de contatos por regional (Para)
    contatos_para = {
        "Campinas": "flavia.pinho@mrv.com.br; ana.tilli@mrv.com.br",
        "Ribeirão Preto": "kaylana.alves@mrv.com.br",
        "Centro Oeste": "nicole.souza@mrv.com.br; maksuel.araujo@mrv.com.br; eunice.prudente@primeconstrucoes.com.br; maryanne.camargo@primeconstrucoes.com.br",
        "Nordeste": "langela.santos@mrv.com.br",
        "Sul": "victoria.gomes@mrv.com.br; filipe.avila@mrv.com.br; simone.csantos@mrv.com.br; monique.silva@mrv.com.br",
        "São Paulo": "telma.amattos@mrv.com.br; cristina.demetrio@parceiro.mrv.com.br; manoella.camargo@mrv.com.br; luciano.lsilva@mrv.com.br; nicoli.santos@mrv.com.br",
        "Triângulo": "kamilly.silva@mrv.com.br; kaylana.alves@mrv.com.br; maria.fernnanda@mrv.com.br"
    }

    # Configurações de tempo (Saudação e Prazo)
    agora = datetime.now()
    saudacao = "Bom dia" if agora.hour < 12 else "Boa tarde"
    
    prazo_rateio = agora + timedelta(hours=32)
    prazo_formatado = prazo_rateio.strftime("%d/%m/%Y às %H:%M")

    # Corpo do email em HTML (usando <br> para quebras de linha)
    corpo_email = f"""
    <p style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000;">
        {saudacao}, Prezado(s)!<br><br>
        Segue em anexo o extrato dos Correios. O rateio deverá ser enviado até <b>{prazo_formatado}</b>.<br><br>
        Atenciosamente,
    </p>
    """

    # Inicializa o Outlook
    outlook = win32.Dispatch('outlook.application')

    # 3. Entrar em cada uma das pastas (exceto "BH")
    pastas_regionais = os.listdir(caminho_mes_recente)
    
    for regional in pastas_regionais:
        caminho_regional = os.path.join(caminho_mes_recente, regional)
        
        # Ignora se for arquivo ou se for a pasta "BH"
        if not os.path.isdir(caminho_regional) or regional.upper() == "BH":
            continue

        print(f"Gerando rascunho para: {regional}...")

        # Cria o email
        mail = outlook.CreateItem(0)
        
        # 6. Preenche o "Para" baseado no dicionário
        mail.To = contatos_para.get(regional, "")
        
        # 5. Preenche o "CC"
        cc_padrao = "vanessa.brodrigues@mrv.com.br; correiosbh@mrv.com.br"
        if regional in ["Triângulo", "Ribeirão Preto"]:
            mail.CC = f"{cc_padrao}; conceicao@mrv.com.br"
        else:
            mail.CC = cc_padrao

        # Assunto
        mail.Subject = f"RES: Extrato Correios - {regional} ({nome_mes})"

        # 4. Anexar arquivos da pasta
        arquivos_na_pasta = os.listdir(caminho_regional)
        for arquivo in arquivos_na_pasta:
            caminho_arquivo = os.path.join(caminho_regional, arquivo)
            if os.path.isfile(caminho_arquivo):
                mail.Attachments.Add(caminho_arquivo)

        # Truque para manter a assinatura padrão do Outlook (Imagem 2)
        # O Display() gera o HTML da assinatura configurada no seu Outlook
        mail.Display() 
        assinatura_outlook = mail.HTMLBody
        
        # Injeta o corpo do texto antes da assinatura
        mail.HTMLBody = f"""
        <html>
        <body>
            {corpo_email}
            {assinatura_outlook}
        </body>
        </html>
        """

        # Salva como rascunho e fecha a janela que o Display() abriu
        mail.Save()
        mail.Close(0) # 0 = olDiscard (fecha a janela sem perguntar, pois já salvamos)

    print("\nProcesso concluído! Verifique a pasta 'Rascunhos' no seu Outlook.")

# Executa a função
if __name__ == "__main__":
    criar_rascunhos_correios()
