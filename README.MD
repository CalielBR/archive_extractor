# 📦 Extração de Arquivos (.zip, .7z, .rar, .zip.001...) com Interface Gráfica (Tkinter)

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

Este script oferece uma interface gráfica (GUI) para selecionar arquivos compactados  
(.zip, .rar, .7z, e arquivos divididos como .zip.001, .zip.002, etc.) e extrair todos  
automaticamente para uma pasta de destino definida pelo usuário.

Ele exibe o progresso da extração em uma barra visual, incluindo o tamanho dos arquivos  
em GB e o número total de arquivos sendo processados.

---

## 🔧 FUNCIONALIDADES

- ✔️ Interface gráfica (GUI) com Tkinter  
- ✔️ Suporte aos formatos `.zip`, `.rar`, `.7z` e `.zip.001/.zip.002/...`  
- ✔️ Junta automaticamente partes divididas (`.zip.001`, `.zip.002`, ...)  
- ✔️ Exibe barra de progresso com GB extraído e total  
- ✔️ Janela com contagem regressiva para fechamento automático após concluir  
- ✔️ Evita fechar caso nenhum arquivo tenha sido selecionado  
- ✔️ Suporte completo para **Windows**, **Linux** e **macOS**  

---

## 📁 REQUISITOS

- ✔️ Python 3.7 ou superior  
- ✔️ Bibliotecas utilizadas (instale com `pip`):  
  - `tkinter` → interface gráfica (já vem com Python)  
  - `zipfile` → módulo padrão do Python  
  - `rarfile` → para extração de arquivos `.rar`  
  - `py7zr` → para extração de arquivos `.7z`  
  - `python-magic-bin` (opcional) → detecção mais precisa de tipos de arquivos  

---

## 💻 INSTALAÇÃO DO PYTHON

### 🪟 Windows:
1. Baixe o instalador: https://www.python.org/downloads/windows/  
2. Marque a opção **"Add Python to PATH"** na instalação  
3. Conclua a instalação normalmente  

### 🐧 Linux (Ubuntu, Mint, etc):
Verifique se o Python já está instalado:
```bash
python3 --version

