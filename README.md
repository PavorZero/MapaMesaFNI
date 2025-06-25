# 🪑 Mapa de Mesa - Gerador de Layout Automatizado

![Python](https://img.shields.io/badge/python-3.6%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## 📌 Visão Geral

Script Python que gera automaticamente um mapa visual de assentos em formato PNG a partir de uma planilha Excel, ideal para eventos com disposição de mesa retangular e 100 participantes.

## ✨ Features

- **Conversão automática** de dados Excel para mapa visual
- **Layout organizado** com 100 assentos:
  - 5 cadeiras superiores (C01-C05)
  - 5 cadeiras inferiores (C06-C10)
  - 90 cadeiras laterais (45 de cada lado)
- **Visualização clara** com número da cadeira e nome do participante
- **Saída em alta resolução** (PNG 150dpi)

## 📦 Pré-requisitos

```bash
# Instale as dependências
pip install pandas matplotlib openpyxl
