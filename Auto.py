import pyautogui as pa
import time
import datetime
import os
import pandas as pd
from tkinter import Tk, Label, Button, PhotoImage
from PIL import Image, ImageTk


pa.PAUSE = 1

# Função para fechar a janela
def finalizar_janela():
    janela.destroy()

# verificar se o login foi esta corretot

def verificar_login_sucesso():
    # Espera até 5 segundos para a página de especifica
    for _ in range(5):
        time.sleep(1)  # Aguarda 1 segundo
        # Verifica se a imagem específica está correta na pasta
        if pa.locateOnScreen('C:/Users/DouglasCavalcante/Desktop/image/login1.png', confidence=0.9):
            return True
    return False

# Abrir Redesoft
pa.hotkey('win', 'd')

pa.click(x=52, y=517)
pa.press('enter')
time.sleep(15) # Espera até 15 segundos para a página de login para digitar usuario e senha
pa.write('DOUGLAS.CAVALCANTE')
pa.press('tab')
pa.write('123')
pa.press('tab')
pa.press('tab')
pa.press('enter')
time.sleep(60) # Espera até 60 segundos para o B2click abrir


# Verifica se o login foi bem-sucedido
if verificar_login_sucesso():
    print("Login bem-sucedido!")

pa.press('enter')
time.sleep(5)

# Abrindo Relatorio mes
pa.hotkey('ctrl', 'f')
pa.write('Vendas por vendedor (APP)')
pa.press('enter')
pa.press('enter')
pa.press('tab')
pa.press('tab')
pa.press('tab')
pa.press('tab')
pa.press('n')
pa.press('tab')
pa.press('enter')
pa.click(x=669, y=642)
pa.press('enter')
time.sleep(50)

#Tratando Relatorio mes

pa.click(x=25, y=279)
pa.click(button='right')
pa.write('e')
pa.write('e')
pa.press('enter')

#Convertendo Relatorio em Tabela mes
pa.click(x=109, y=275)
pa.press('alt')
pa.press('c')
pa.press('t')
pa.press('enter')
pa.press('enter')

#Nomeando Tabela
pa.click(x=130, y=100)
pa.write('Vendas_Mes')
pa.press('enter')

#Salvar OneDrive
pa.press('alt')
pa.press('a')
pa.press('r')
pa.press('2')
time.sleep(1)
pa.click(x=502, y=491)
time.sleep(1)
pa.click(x=298, y=514)
time.sleep(1)
pa.click(x=437, y=458)
pa.write('Vendas por vendedor (APP)')
pa.press('enter')
pa.click(x=1029, y=531)
pa.press('enter')
pa.hotkey('alt', 'f4')

#Abrindo Relatorio do dia
time.sleep(5)
data_atual = datetime.date.today()
data = data_atual.strftime('%d%m%Y')
pa.write(data)
pa.press('tab')
data_atual = datetime.date.today()
data = data_atual.strftime('%d%m%Y')
pa.write(data)
pa.press('tab')
pa.press('tab')
pa.press('tab')
pa.press('n')
pa.press('tab')
pa.press('enter')
pa.click(x=669, y=642)
pa.press('enter')
time.sleep(25)

#Tratando Relatorio
pa.click(x=25, y=279)
pa.click(button='right')
pa.write('e')
pa.write('e')
pa.press('enter')

#Convertendo Relatorio em Tabela dia
pa.click(x=109, y=275)
pa.press('alt')
pa.press('c')
pa.press('t')
pa.press('enter')
pa.press('enter')

#Salvar OneDrive
pa.press('alt')
pa.press('a')
pa.press('r')
pa.press('2')
pa.click(x=502, y=491)
pa.click(x=298, y=514)
pa.click(x=437, y=458)
pa.write('Vendas por vendedor (APP) Dia')
pa.press('enter')
pa.click(x=1029, y=531)
pa.press('enter')
pa.hotkey('alt', 'f4')

#Fechando aba de Vendas
pa.hotkey('alt', 'f4')

#Fechando Redesoft
pa.hotkey('alt', 'f4')

# Caminhos dos arquivos Excel
excel_file_path1 = 'C:/Users/DouglasCavalcante/OneDrive - Grupo Pichilau/Aplicativo_Base/Vendas por vendedor (APP).xlsx'
excel_file_path2 = 'C:/Users/DouglasCavalcante/OneDrive - Grupo Pichilau/Aplicativo_Base/Vendas por vendedor (APP) Dia.xlsx'

# Ler os arquivos Excel
df1 = pd.read_excel(excel_file_path1, engine='openpyxl')
df2 = pd.read_excel(excel_file_path2, engine='openpyxl')

# Caminhos para salvar os arquivos CSV
csv_file_path1 = 'C:/Users/DouglasCavalcante/Desktop/APPvendasOnline/vendaspichilau/vendas.csv'
csv_file_path2 = 'C:/Users/DouglasCavalcante/Desktop/APPvendasOnline/vendaspichilau/vendas_dia.csv'

# Salvar os DataFrames como CSV
df1.to_csv(csv_file_path1, index=False)
df2.to_csv(csv_file_path2, index=False)

print("Arquivos CSV foram criados com sucesso!")


#Abrindo Whatsapp ja logado no Computador

time.sleep(2)
pa.press('win')
pa.write('whatsapp')
time.sleep(1)
pa.press('enter')
time.sleep(1)
pa.hotkey('ctrl', 'f')
time.sleep(5)

#Buscando para quem ira enviar msg de confirmação de vendas
pa.write('Douglas Camelo')
time.sleep(3)
pa.click(x=247, y=230)
pa.write('Vendas Atualizadas com Sucesso!!')
pa.press('enter')
time.sleep(3)
pa.hotkey('alt', 'f4')

def finalizado():
    hora_atual = datetime.datetime.now().strftime("%H:%M:%S")
    label.config(text=f"Vendas atualizadas com sucesso às {hora_atual}")
    # Agendar o fechamento da janela após 5 minutos
    janela.after(300000, finalizar_janela)  # 300000 milissegundos = 5 minutos

# Criar a janela
janela = Tk()
janela.title("Vendas Atualizadas")

# Definir o tamanho da janela
largura_janela = 600
altura_janela = 400

# Obter as dimensões da tela
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()

# Calcular as coordenadas para centralizar a janela na tela
x = (largura_tela - largura_janela) // 2
y = (altura_tela - altura_janela) // 2

# Definir o tamanho e a posição da janela
janela.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")

# Definir a cor de fundo como laranja
janela.configure(bg='orange')

# Adicionar um rótulo
label = Label(janela, text="Vendas atualizadas com sucesso")
label.pack(pady=20)

# Adicionar um botão para fechar a janela
botao_fechar = Button(janela, text="Fechar", command=janela.destroy)
botao_fechar.pack(pady=10)

# Carregar a imagem da logo
caminho_logo = "C:/Users/DouglasCavalcante/Desktop/logo.png"  # Substitua pelo caminho da sua imagem
logo_pil = Image.open(caminho_logo)
largura_logo, altura_logo = botao_fechar.winfo_width(), botao_fechar.winfo_height()
logo_pil = logo_pil.resize((75, 75), Image.LANCZOS)  # Redimensionar a imagem
logo = ImageTk.PhotoImage(logo_pil)
logo_label = Label(janela, image=logo, bg='orange')
logo_label.pack(pady=20)

# Chamando a função 'finalizado'
finalizado()

# Executar a janela
janela.mainloop()

