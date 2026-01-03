import discord
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from flask import Flask
from threading import Thread
import os

app = Flask('')

@app.route('/')
def home():
    return "Bot Angelotti está rodando!"

def run():
    app.run(host='0.0.0.0', port=5000)

def keep_alive():
    t = Thread(target=run)
    t.start()

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]

json_file = "ControleAngelotti.json"
if os.path.exists(json_file):
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_file, scope)
    client_sheets = gspread.authorize(creds)
else:
    print(f"❌ Erro: Arquivo {json_file} não encontrado!")
    client_sheets = None

SPREADSHEET_ID = "1gCnfcx4BMyqpBlM3gSLErEcOAJ6KTiOAfRZhLYitumM"

INTENTS = discord.Intents.default()
INTENTS.messages = True
INTENTS.message_content = True
client_discord = discord.Client(intents=INTENTS)

@client_discord.event
async def on_ready():
    print(f'✅ Bot conectado como {client_discord.user}')
    if client_sheets:
        try:
            spreadsheet = client_sheets.open_by_key(SPREADSHEET_ID)
            abas = [ws.title for ws in spreadsheet.worksheets()]
            print("📄 Abas disponíveis:", abas)
        except Exception as e:
            print("❌ Erro ao acessar a planilha:", e)

@client_discord.event
async def on_message(message):
    if message.author == client_discord.user:
        return

    conteudo = message.content.strip()
    data_msg = message.created_at.strftime("%d/%m/%Y")

    if not client_sheets:
        if conteudo.startswith('!'):
            await message.channel.send("❌ Erro: Credenciais do Google Sheets não configuradas no servidor.")
        return

    if conteudo.startswith('!testar'):
        try:
            spreadsheet = client_sheets.open_by_key(SPREADSHEET_ID)
            abas = [ws.title for ws in spreadsheet.worksheets()]
            await message.channel.send(f"✅ A planilha está acessível. Abas encontradas: {abas}")
        except Exception as e:
            await message.channel.send(f"❌ Erro ao acessar a planilha: {e}")
        return

    if conteudo.startswith('!NovoSKU') and "Licença:" in conteudo:
        linhas = [l.strip() for l in conteudo.splitlines() if l.strip()]
        if len(linhas) < 4:
            await message.channel.send(
                "❌ Formato incompleto. Envie 4 linhas:\n"
                "1️⃣ !NovoSKU Licença: [nome da aba]\n"
                "2️⃣ Assunto do projeto\n"
                "3️⃣ Nome do licenciado\n"
                "4️⃣ Código Angelotti"
            )
            return

        licenca = linhas[0].split("Licença:")[1].strip()
        assunto = linhas[1]
        licenciado = linhas[2]
        codigo = linhas[3]

        try:
            spreadsheet = client_sheets.open_by_key(SPREADSHEET_ID)
            sheet = spreadsheet.worksheet(licenca)

            if sheet.find(codigo):
                await message.channel.send(f"⚠️ O código **{codigo}** já existe na aba **{licenca}**.")
                return

            valores = [
                codigo,
                licenciado,
                assunto,
                f"ENVIADO {data_msg}",
                "Aguardando amostra",
                "Não"
            ]
            sheet.append_row(valores)
            await message.channel.send(f'✅ Dados registrados na aba **{licenca}**!')
        except gspread.exceptions.WorksheetNotFound:
            await message.channel.send(f"❌ Aba **{licenca}** não encontrada.")
        except Exception as e:
            await message.channel.send(f"❌ Erro ao processar: {e}")
        return

    elif conteudo.startswith('!NovoSKU'):
        await message.channel.send("❌ Mensagem não contém 'Licença:' para identificar a aba.")
        return

    async def atualizar_status(comando, coluna, texto, sobrescrever=False):
        try:
            linhas = [l.strip() for l in conteudo.splitlines() if l.strip()]
            if len(linhas) < 2 or "Licença:" not in linhas[0]:
                await message.channel.send("❌ Formato inválido. Use duas linhas:\n1️⃣ Comando com 'Licença: NomeDaAba'\n2️⃣ Código Angelotti")
                return

            licenca = linhas[0].split("Licença:")[1].strip()
            codigo = linhas[1]

            spreadsheet = client_sheets.open_by_key(SPREADSHEET_ID)
            sheet = spreadsheet.worksheet(licenca)

            cell = sheet.find(codigo)
            if cell:
                linha = cell.row
                if sobrescrever:
                    novo_valor = f"{texto} {data_msg}"
                else:
                    valor_atual = sheet.cell(linha, coluna).value or ""
                    novo_valor = f"{valor_atual}\n{texto} {data_msg}".strip()

                sheet.update_cell(linha, coluna, novo_valor)
                await message.channel.send(f'✅ Atualização feita para **{codigo}** na aba **{licenca}**.')
            else:
                await message.channel.send(f"❌ Código **{codigo}** não encontrado na aba **{licenca}**.")
        except Exception as e:
            await message.channel.send(f"❌ Erro: {e}")

    if conteudo.startswith('!AprovadoConceito'):
        await atualizar_status('!AprovadoConceito', 4, "APROVADO")

    if conteudo.startswith('!RevisãoConceito'):
        await atualizar_status('!RevisãoConceito', 4, "PEDIDO DE REVISÃO")

    if conteudo.startswith('!EnvioAmostra'): 
        await atualizar_status('!EnvioAmostra', 5, "ENVIADO", sobrescrever=True)
            
    if conteudo.startswith('!AprovadaAmostra'):
        await atualizar_status('!AprovadaAmostra', 5, "APROVADO")
    
    if conteudo.startswith('!RevisãoAmostra'):
        await atualizar_status('!RevisãoAmostra', 5, "PEDIDO DE REVISÃO")

if __name__ == "__main__":
    keep_alive()
    TOKEN = os.getenv("DISCORD_TOKEN")
    if not TOKEN:
        print("❌ Erro: DISCORD_TOKEN não encontrado nos Secrets!")
    else:
        try:
            client_discord.run(TOKEN)
        except Exception as e:
            print(f"❌ Erro ao iniciar o bot: {e}")
