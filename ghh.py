import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Abrir o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
# Aguardar o carregamento completo (ajuste o tempo conforme necessário)
sleep(20)

# Ler Planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('Pasta12.xlsx')
pagina_clientes = workbook["Planilha1"]

for linha in pagina_clientes.iter_rows(min_row=2):
    # Extrair informações da linha
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    valorboletop = linha[3].value 

    print(nome)
    print(telefone)
    print(vencimento)
    print(valorboletop)
    
    # Construir a mensagem personalizada
    mensagem = f'Prezado {nome}, Gostaríamos de lembrá-lo(a) de que o pagamento referente à sua fatura no valor de R$ {valorboletop} está pendente. A data de vencimento original era {vencimento.strftime("%d/%m/%Y")} e, até o momento, não registramos o recebimento. Favor pagar no link https://mpago.la/1EL9NTP'

    # Criar link personalizado do WhatsApp
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    
    try:
        # Abrir o link do WhatsApp para enviar a mensagem
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)  # Aguardar o carregamento da página de mensagem

        # Localizar e clicar na seta para enviar a mensagem
        seta = pyautogui.locateCenterOnScreen('seta.png')
        if seta is not None:
            pyautogui.click(seta[0], seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')  # Fechar a aba
            sleep(5)
        else:
            raise Exception('Não foi possível localizar a seta para enviar a mensagem')
    
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {e}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone},{os.linesep}')