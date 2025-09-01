"""
═══════════════════════════════════════════════════════════════════════════════
📦 SCRIPT: Extração de Arquivos (.zip, .7z, .rar, .zip.001...) com Interface Gráfica (Tkinter)

Este script oferece uma interface gráfica (GUI) para selecionar arquivos compactados
(.zip, .rar, .7z, e arquivos divididos como .zip.001, .zip.002, etc.) e extrair todos
automaticamente para uma pasta de destino definida pelo usuário.

Ele exibe o progresso da extração em uma barra visual, incluindo o tamanho dos arquivos
em GB e o número total de arquivos sendo processados.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🔧 FUNCIONALIDADES:

✔️ Interface gráfica (GUI) com Tkinter  
✔️ Suporte aos formatos .zip, .rar, .7z e .zip.001/.zip.002/...  
✔️ Junta automaticamente partes divididas (.zip.001, .zip.002, ...)  
✔️ Exibe barra de progresso com GB extraído e total  
✔️ Janela com contagem regressiva para fechamento automático após concluir  
✔️ Evita fechar caso nenhum arquivo tenha sido selecionado  
✔️ Suporte completo para Windows, Linux e macOS  

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📁 REQUISITOS:

✔️ Python 3.7 ou superior  
✔️ Bibliotecas usadas (instale com pip):  
    - tkinter (interface gráfica) → geralmente já vem no Python  
    - zipfile (padrão do Python)  
    - rarfile → extração de .rar  
    - py7zr   → extração de .7z  
    - python-magic-bin (opcional) → melhor detecção de tipos de arquivos  

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
💻 INSTALAÇÃO DO PYTHON:

🔹 WINDOWS:  
    1. Baixe o instalador: https://www.python.org/downloads/windows/  
    2. Marque a opção "Add Python to PATH" na instalação  
    3. Conclua a instalação normalmente  

🔹 LINUX (Ubuntu, Mint, etc):  
    Verifique se o Python já está instalado com:  
        python3 --version  
    Caso não esteja ou deseje atualizar:  
        sudo apt update  
        sudo apt install python3 python3-pip python3-tk unrar  

🔹 MACOS:  
    Python já vem instalado, mas você pode atualizar usando Homebrew:  
        brew install python  
        brew install unrar  

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📦 INSTALAÇÃO DAS DEPENDÊNCIAS:

Abra o terminal ou prompt de comando e execute:

    pip install py7zr rarfile python-magic-bin

Caso use múltiplas versões do Python, use:

    python3 -m pip install py7zr rarfile python-magic-bin

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
💻 EDITOR RECOMENDADO:

👉 VS Code (Visual Studio Code):  
Baixe e instale em: https://code.visualstudio.com/  

No VS Code, abra seu script e use o terminal integrado para executar.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
▶️ COMO EXECUTAR O SCRIPT:

1. Salve este script como `unzipfull.py`  
2. Abra o terminal ou prompt de comando na pasta onde está o arquivo  
3. Execute o script com:

       python unzipfull.py
    ou
       python3 unzipfull.py

4. A interface será aberta. Selecione os arquivos e a pasta de destino  
5. Acompanhe a extração na barra de progresso e na área de mensagens  

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🚨 NOTAS IMPORTANTES:

- O `rarfile` precisa do programa `unrar` instalado no sistema (Linux/macOS)  
- Em sistemas Linux, o `tkinter` pode precisar ser instalado separadamente com:
      sudo apt install python3-tk

- Evite mexer nos arquivos enquanto estão sendo extraídos  
- Ao final, a janela exibirá uma contagem regressiva antes de fechar  

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📞 DÚVIDAS?  
Se tiver dúvidas, revise a documentação de cada biblioteca:
- https://pypi.org/project/py7zr/  
- https://pypi.org/project/rarfile/  
- https://docs.python.org/3/library/tk.html  

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
👨‍💻 CRIADO POR:
Mike CalielFR Barbosa  
https://github.com/Calielbr  
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

"""


import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import time
import zipfile
import rarfile
import py7zr


class Status:
    """Armazena o progresso da extração."""

    def __init__(self):
        self.total_arquivos = 0
        self.total_gb_arquivos = 0.0
        self.arquivo_atual_index = 0


status = Status()


def juntar_partes_zip(folder: str, base_name: str) -> str | None:
    """
    Junta partes de um arquivo zip dividido em um único arquivo.

    Args:
        folder: Diretório contendo os segmentos .zip.xxx.
        base_name: Nome-base do arquivo (antes do .zip.xxx).

    Returns:
        Caminho do arquivo zip unido, ou None se não encontrar partes.
    """
    partes = []
    i = 1
    while True:
        parte = os.path.join(folder, f"{base_name}.zip.{i:03d}")
        if not os.path.exists(parte):
            break
        partes.append(parte)
        i += 1

    if not partes:
        return None

    unido_path = os.path.join(folder, f"{base_name}_unido.zip")
    with open(unido_path, "wb") as outfile:
        for parte in partes:
            with open(parte, "rb") as infile:
                outfile.write(infile.read())
    return unido_path


def extract_archive(archive_path: str, destination_folder: str, update_callback=None) -> str:
    """
    Extrai um arquivo compactado (zip, rar, 7z ou partes zip).

    Args:
        archive_path: Caminho do arquivo de entrada.
        destination_folder: Pasta de destino para os arquivos extraídos.
        update_callback: Função chamada após cada extração interna para atualizar o progresso.

    Returns:
        Mensagem de sucesso ou erro da extração.
    """
    tamanho_bytes = os.path.getsize(archive_path)
    tamanho_gb = tamanho_bytes / (1024 ** 3)

    try:
        lower = archive_path.lower()
        if lower.endswith(".zip"):
            zip_ref = zipfile.ZipFile(archive_path, "r")
            namelist = zip_ref.namelist()
            total_interno = len(namelist)
            for i, name in enumerate(namelist, 1):
                zip_ref.extract(name, destination_folder)
                if update_callback:
                    update_callback(i, total_interno, tamanho_gb)
            zip_ref.close()
        elif lower.endswith(".rar"):
            rar_ref = rarfile.RarFile(archive_path, "r")
            namelist = rar_ref.namelist()
            total_interno = len(namelist)
            for i, name in enumerate(namelist, 1):
                rar_ref.extract(name, destination_folder)
                if update_callback:
                    update_callback(i, total_interno, tamanho_gb)
            rar_ref.close()
        elif lower.endswith(".7z"):
            seven_ref = py7zr.SevenZipFile(archive_path, "r")
            allnames = seven_ref.getnames()
            total_interno = len(allnames)
            for i, name in enumerate(allnames, 1):
                seven_ref.extract(targets=[name], path=destination_folder)
                if update_callback:
                    update_callback(i, total_interno, tamanho_gb)
            seven_ref.close()
        elif ".zip." in lower:
            folder = os.path.dirname(archive_path)
            base = os.path.basename(archive_path).split(".zip.")[0]
            unido = juntar_partes_zip(folder, base)
            if unido:
                zip_ref = zipfile.ZipFile(unido, "r")
                namelist = zip_ref.namelist()
                total_interno = len(namelist)
                for i, name in enumerate(namelist, 1):
                    zip_ref.extract(name, destination_folder)
                    if update_callback:
                        update_callback(i, total_interno, tamanho_gb)
                zip_ref.close()
                os.remove(unido)
            else:
                return (f"❌ Partes não encontradas: {archive_path} "
                        f"({tamanho_gb:.2f} GB)")
        else:
            return (f"Formato não suportado: {os.path.basename(archive_path)} "
                    f"({tamanho_gb:.2f} GB)")

        return (f"✅ Extraído com sucesso: {os.path.basename(archive_path)} "
                f"({tamanho_gb:.2f} GB)")
    except (zipfile.BadZipFile, rarfile.Error, py7zr.exceptions.Bad7zFile) as e:
        return (f"❌ Erro ao extrair {os.path.basename(archive_path)} "
                f"({tamanho_gb:.2f} GB): {e}")


def escolher_arquivos():
    """Abre diálogo para selecionar arquivos compactados e popula o Listbox com os paths."""
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos para extrair",
        filetypes=[("Arquivos compactados",
                    "*.zip *.rar *.7z *.001 *.002 *.003")],
    )
    lista_arquivos.delete(0, tk.END)
    for arq in arquivos:
        lista_arquivos.insert(tk.END, arq)


def escolher_pasta():
    """Abre diálogo para seleção da pasta de destino e atualiza o campo correspondente."""
    pasta = filedialog.askdirectory(title="Selecione a pasta de destino")
    if pasta:
        destino_var.set(pasta)


def fechar_apos_contagem(parent: tk.Tk | tk.Toplevel,
                         segundos: int = 60,
                         mensagem: str | None = None):
    """
    Exibe uma janela de aviso com contagem regressiva para fechamento.

    Args:
        parent: Janela que será fechada após a contagem.
        segundos: Tempo em segundos da contagem regressiva.
        mensagem: Mensagem exibida na janela.
    """
    def countdown():
        for sec in range(segundos, 0, -1):
            label_contagem.config(text=f"Fechando em {sec} segundos…")
            time.sleep(1)
        parent.quit()

    janela = tk.Toplevel(parent)
    janela.title("Aviso")
    janela.geometry("350x120")
    janela.protocol("WM_DELETE_WINDOW", parent.quit)
    tk.Label(janela, text=mensagem or "Nenhum arquivo para extração.",
             font=("Arial", 12)).pack(pady=10)
    label_contagem = tk.Label(janela, text=f"Fechando em {segundos} segundos…",
                              font=("Arial", 10))
    label_contagem.pack()
    threading.Thread(target=countdown, daemon=True).start()


def atualizar_progresso_atual(i: int,
                              total_interno: int,
                              tamanho_gb: float):
    """
    Atualiza a barra de progresso e o texto dinâmico com o progresso atual.

    Args:
        i: Índice do próximo arquivo interno a extrair.
        total_interno: Quantidade total de arquivos no arquivamento.
        tamanho_gb: Tamanho em GB do arquivo em extração atual.
    """
    progresso['maximum'] = total_interno
    progresso['value'] = i
    parcial = (i / total_interno) * tamanho_gb
    label_status.config(
        text=(
            f"Extraindo arquivo {status.arquivo_atual_index} de "
            f"{status.total_arquivos} ({parcial:.2f} GB extraído no "
            f"momento de {status.total_gb_arquivos:.2f} GB total)"
        )
    )


def extrair_todos():
    """
    Controla o fluxo da extração, atualiza estatísticas gerais e inicia a thread.
    """
    destino = destino_var.get()
    if not destino:
        messagebox.showwarning("Atenção", "Selecione uma pasta de destino.")
        return

    arquivos = lista_arquivos.get(0, tk.END)
    if not arquivos:
        fechar_apos_contagem(root, 60)
        return

    status.total_arquivos = len(arquivos)
    status.total_gb_arquivos = sum(
        os.path.getsize(arq) / (1024 ** 3) for arq in arquivos
    )

    def tarefa():
        resultados.delete(1.0, tk.END)
        for idx, arq in enumerate(arquivos, start=1):
            status.arquivo_atual_index = idx
            msg = extract_archive(arq, destino,
                                  update_callback=atualizar_progresso_atual)
            resultados.insert(tk.END, msg + "\n")

        label_status.config(
            text=(
                f"✅ Extração concluída: {status.total_arquivos} "
                f"arquivo(s) / {status.total_gb_arquivos:.2f} GB"
            )
        )
        fechar_apos_contagem(
            root,
            60,
            mensagem="✅ Extração concluída. A janela fechará automaticamente."
        )
        progresso['value'] = 0

    threading.Thread(target=tarefa, daemon=True).start()


# --- Interface Gráfica ---
root = tk.Tk()
root.title("Extrator de Arquivos")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

tk.Button(frame, text="Selecionar Arquivos", command=escolher_arquivos)\
  .grid(row=0, column=0, sticky="ew")
lista_arquivos = tk.Listbox(frame, width=80, height=10)
lista_arquivos.grid(row=1, column=0, columnspan=3, pady=5)

tk.Button(frame, text="Selecionar Pasta Destino", command=escolher_pasta)\
  .grid(row=2, column=0, sticky="ew", pady=5)
destino_var = tk.StringVar()
tk.Entry(frame, textvariable=destino_var, width=60)\
  .grid(row=2, column=1, sticky="ew")
tk.Button(frame, text="Extrair", command=extrair_todos)\
  .grid(row=2, column=2, sticky="ew", padx=5)

progresso = ttk.Progressbar(frame, length=500)
progresso.grid(row=4, column=0, columnspan=3, pady=5)

label_status = tk.Label(frame, text="Status: aguardando extração…")
label_status.grid(row=5, column=0, columnspan=3)

resultados = scrolledtext.ScrolledText(frame, width=80, height=10)
resultados.grid(row=6, column=0, columnspan=3, pady=10)

root.mainloop()
