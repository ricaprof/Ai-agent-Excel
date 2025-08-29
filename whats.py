import openpyxl
import webbrowser
from urllib.parse import quote

ARQUIVO = "lista_compras.xlsx"

def enviar_lista_whatsapp():
    wb = openpyxl.load_workbook(ARQUIVO)
    ws = wb.active
    produtos = [row[0] for row in ws.iter_rows(min_row=2, values_only=True) if row[0]]
    print(produtos)

    if produtos:
        mensagem = "Lista de compras: " + ", ".join(produtos)
        link = f"https://wa.me/?text={quote(mensagem)}"
        print("[AGENTE] Abrindo WhatsApp...")
        webbrowser.open(link)  # abre no navegador
    else:
        print("[AGENTE] Sua lista está vazia.")

if __name__ == "__main__":
    enviar_lista_whatsapp()
