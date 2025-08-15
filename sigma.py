import os
import time
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from PyPDF2 import PdfReader
from reportlab.lib.units import mm
import win32print
import ctypes.wintypes
import traceback
import psutil

USUARIO = os.getlogin()
PASTA_DOWNLOADS = os.path.join(os.environ["USERPROFILE"], "Downloads")
PASTA_WATCH = PASTA_DOWNLOADS  # Agora monitora Downloads direto
PASTA_TEMP = os.path.join(os.environ['TEMP'], 'etiquetas')

os.makedirs(PASTA_TEMP, exist_ok=True)

monitorando = False
largura_etiqueta = 100 * mm
altura_etiqueta = 150 * mm
impressora_selecionada = ""

def fechar_acrobat():
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'] and 'AcroRd32.exe' in proc.info['name']:
                proc.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

def imprimir_pdf(pdf_path):
    try:
        print(f"🖨️ Impressora selecionada: {impressora_selecionada}")
        print(f"📄 Enviando para impressão: {pdf_path}")

        if not os.path.exists(pdf_path):
            print("❌ Arquivo PDF não encontrado!")
            return

        impressoras = [p[2] for p in win32print.EnumPrinters(2)]
        if impressora_selecionada not in impressoras:
            print(f"❌ Impressora '{impressora_selecionada}' não encontrada no sistema!")
            return

        acrobat_path = r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
        if not os.path.exists(acrobat_path):
            print("❌ Adobe Acrobat Reader não encontrado!")
            return

        subprocess.Popen([
            acrobat_path,
            "/t",
            pdf_path,
            impressora_selecionada
        ])

        time.sleep(5)
        fechar_acrobat()

        print("✅ Impressão enviada com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao imprimir com Acrobat: {e}")
        traceback.print_exc()

def monitorar():
    global monitorando
    print(f"📂 Monitorando a pasta: {PASTA_WATCH}")
    arquivos_vistos = set()

    while monitorando:
        time.sleep(2)
        arquivos = [
            f for f in os.listdir(PASTA_WATCH)
            if f.lower().endswith(".pdf") and f.startswith("idRecibo_")
        ]
        novos = set(arquivos) - arquivos_vistos

        for nome_arquivo in novos:
            caminho_pdf = os.path.join(PASTA_WATCH, nome_arquivo)
            print(f"📄 Novo PDF detectado: {nome_arquivo}")
            try:
                reader = PdfReader(caminho_pdf)
                num_pages = len(reader.pages)

                if num_pages == 0:
                    print("⚠️ PDF vazio, será removido.")
                    os.remove(caminho_pdf)
                    continue

                if num_pages == 1:
                    page = reader.pages[0]
                    width = float(page.mediabox.width)
                    height = float(page.mediabox.height)
                    largura_mm = width / mm
                    altura_mm = height / mm
                    tolerancia = 2

                    if abs(largura_mm - 100) <= tolerancia and abs(altura_mm - 150) <= tolerancia:
                        print("📄 PDF com 1 página e tamanho correto detectado. Imprimindo direto...")
                        imprimir_pdf(caminho_pdf)
                        os.remove(caminho_pdf)
                        print(f"🗑️ PDF impresso e removido: {nome_arquivo}")
                        continue

                print(f"📄 PDF com {num_pages} página(s). Imprimindo como está...")
                imprimir_pdf(caminho_pdf)
                os.remove(caminho_pdf)
                print(f"🗑️ PDF impresso e removido: {nome_arquivo}")

            except Exception as e:
                print(f"❌ Erro ao processar PDF {nome_arquivo}: {e}")
                traceback.print_exc()

        arquivos_vistos.update(novos)

def iniciar_monitoramento():
    global largura_etiqueta, altura_etiqueta, impressora_selecionada, monitorando
    try:
        largura = float(entry_largura.get())
        altura = float(entry_altura.get())
        largura_etiqueta = largura * mm
        altura_etiqueta = altura * mm
        impressora_selecionada = combo_impressoras.get()

        if not impressora_selecionada:
            messagebox.showerror("Erro", "Selecione uma impressora.")
            return

        monitorando = True
        btn_iniciar.config(state='disabled')
        btn_parar.config(state='normal')
        threading.Thread(target=monitorar, daemon=True).start()
    except ValueError:
        messagebox.showerror("Erro", "Digite dimensões válidas em milímetros.")
    except Exception as e:
        messagebox.showerror("Erro inesperado", str(e))
        traceback.print_exc()

def parar_monitoramento():
    global monitorando
    monitorando = False
    btn_iniciar.config(state='normal')
    btn_parar.config(state='disabled')
    print("🛑 Monitoramento encerrado.")

# GUI
root = tk.Tk()
root.title("Impressão de Etiquetas")

try:
    tk.Label(root, text="Largura (mm):").grid(row=0, column=0, padx=5, pady=5)
    entry_largura = tk.Entry(root)
    entry_largura.insert(0, "100")
    entry_largura.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(root, text="Altura (mm):").grid(row=1, column=0, padx=5, pady=5)
    entry_altura = tk.Entry(root)
    entry_altura.insert(0, "150")
    entry_altura.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(root, text="Impressora:").grid(row=2, column=0, padx=5, pady=5)
    combo_impressoras = ttk.Combobox(root, values=[printer[2] for printer in win32print.EnumPrinters(2)])
    combo_impressoras.grid(row=2, column=1, padx=5, pady=5)

    btn_iniciar = tk.Button(root, text="Iniciar Monitoramento", command=iniciar_monitoramento)
    btn_iniciar.grid(row=3, column=0, padx=10, pady=10)

    btn_parar = tk.Button(root, text="Parar", state='disabled', command=parar_monitoramento)
    btn_parar.grid(row=3, column=1, padx=10, pady=10)

    root.mainloop()
except Exception as e:
    print(f"Erro ao iniciar a interface: {e}")
    traceback.print_exc()
