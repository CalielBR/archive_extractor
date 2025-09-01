# 📦 Archive Extraction Tool with Graphical Interface (Tkinter)

# unzipfull

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Python Version](https://img.shields.io/badge/Python-3.7%2B-blue.svg)
![GitHub Repo stars](https://img.shields.io/github/stars/CalielBR/unzipfull?style=social)
![GitHub forks](https://img.shields.io/github/forks/CalielBR/unzipfull?style=social)




This script provides a graphical user interface (GUI) for selecting compressed files (.zip, .rar, .7z, and split archives like .zip.001, .zip.002, etc.) and automatically extracting them to a user-defined destination folder.

It displays extraction progress with a visual bar, including file sizes in GB and the total number of files being processed.

## 🔧 Features

✔️ Graphical interface (GUI) with Tkinter  
✔️ Supports .zip, .rar, .7z, and split archives (.zip.001, .zip.002, ...)  
✔️ Automatically combines split archive parts  
✔️ Displays progress bar with GB extracted and total size  
✔️ Countdown window for automatic closure after completion  
✔️ Prevents closing if no files were selected  
✔️ Full support for Windows, Linux, and macOS  

## 📁 Requirements

✔️ Python 3.7 or higher  
✔️ Required libraries (install with pip):  
    - tkinter (graphical interface) → usually included with Python  
    - zipfile (Python standard library)  
    - rarfile → .rar extraction  
    - py7zr → .7z extraction  
    - python-magic-bin (optional) → better file type detection  

## 💻 Python Installation

🔹 **WINDOWS:**  
    1. Download installer: https://www.python.org/downloads/windows/  
    2. Check "Add Python to PATH" during installation  
    3. Complete installation normally  

🔹 **LINUX (Ubuntu, Mint, etc):**  
    Check if Python is already installed:  
        `python3 --version`  
    If not installed or to update:  
        `sudo apt update`  
        `sudo apt install python3 python3-pip python3-tk unrar`  

🔹 **MACOS:**  
    Python comes pre-installed, but you can update using Homebrew:  
        `brew install python`  
        `brew install unrar`  

## 📦 Installing Dependencies

Open terminal or command prompt and execute:

    pip install py7zr rarfile python-magic-bin

If using multiple Python versions, use:

    python3 -m pip install py7zr rarfile python-magic-bin

## 💻 Recommended Editor

👉 **VS Code (Visual Studio Code):**  
Download and install from: https://code.visualstudio.com/  

In VS Code, open your script and use the integrated terminal to execute it.

## ▶️ How to Run the Script

1. Save this script as `unzipfull.py`  
2. Open terminal or command prompt in the script's folder  
3. Execute the script with:

       python unzipfull.py
    or
       python3 unzipfull.py

4. The interface will open. Select files and destination folder  
5. Monitor extraction progress in the progress bar and message area  

## 🚨 Important Notes

- `rarfile` requires the `unrar` program installed on the system (Linux/macOS)  
- On Linux systems, `tkinter` may need separate installation:
      `sudo apt install python3-tk`

- Avoid modifying files while extraction is in progress  
- Upon completion, the window will display a countdown before closing  

## 📞 Need Help?  
If you have questions, review each library's documentation:
- https://pypi.org/project/py7zr/  
- https://pypi.org/project/rarfile/  
- https://docs.python.org/3/library/tk.html  

---
👨‍💻 **Created by:** Mike CalielFR Barbosa  
https://github.com/Calielbr  

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
