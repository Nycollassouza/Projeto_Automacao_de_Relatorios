# Automatização de Relatórios Financeiros

## Problema
Processo manual de extração de comprovantes de portal corporativo > edição Excel > importação base diária levava 10-15min/dia, com erros frequentes em fórmulas/datas.

## Solução
Script Python end-to-end:
- Selenium para login + filtros + download.
- Pandas para ETL (limpeza, mapeamento, derivação colunas).
- OpenPyXL/PyAutoGUI para merge + fórmulas.
- Win32 para recálculo batch.

**Resultados**: Tempo <2min, zero erros, escalável para multi-usuário.

## Como rodar
1. `pip install -r requirements.txt`
2. Edite config (URL, login).
3. `python main.py`


