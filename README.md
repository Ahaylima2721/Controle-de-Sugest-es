# Controle de Sugestões

Aplicativo Python para leitura de período de vendas e previsibilidade (previsão de falta de peças, consulta de vendas por filial, sugestões de compra e transferência).

> Este repositório contém **apenas o código**. Relatórios do ERP (1174, 1211, Vendas Perdidas), a base consolidada (`base/`), as saídas (`saida/`) e o `.exe` gerado **não são versionados** (ver `.gitignore`).

## Requisitos
- Windows
- Python 3.12+ (testado com 3.14) com tkinter (incluso no instalador oficial)

```powershell
python -m pip install -r requirements.txt
```

## Executar
- **Aplicativo de previsão (navegador):** `INICIAR.bat` → abre em `http://localhost:8501`.
- **Atualizar sem navegador:** colocar os relatórios em `entrada\` e rodar `ATUALIZAR.bat`.

As pastas `base\`, `entrada\` e `saida\` são criadas/preenchidas na primeira importação.

## Gerar o `Consulta Vendas.exe`
```powershell
sistema\gerar_exe_consulta.bat
```
Usa PyInstaller (`--onefile --windowed`) e grava `Consulta Vendas.exe` na raiz do projeto. O .exe lê os relatórios que estiverem na mesma pasta dele.

## Manual de uso
Ver [LEIA-ME.md](LEIA-ME.md).

## Estrutura do código
| Arquivo | Função |
|---|---|
| `sistema/app.py` | Servidor local + interface web (`sistema/web/app.html`) |
| `sistema/atualizar.py` | Importação dos relatórios e atualização da base |
| `sistema/leitores.py` | Leitura dos relatórios 1174 / 1211 / Vendas Perdidas |
| `sistema/modelo.py` | Modelo de previsão (backtest, perfis de demanda) |
| `sistema/motor.py` | Cálculo de situação, cobertura e sugestão de compra |
| `sistema/rede_estoque.py` | Transferências e saúde do estoque da rede |
| `sistema/relatorio.py`, `sistema/painel.py` | Excel e painel HTML offline |
| `sistema/consulta_vendas.py` | Programa desktop de consulta (vira o .exe) |
| `sistema/config.py` | Parâmetros de negócio (lead time, segurança, filiais) |
