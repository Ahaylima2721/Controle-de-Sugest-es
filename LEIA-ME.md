# Previsão de Falta de Peças — BETA

## Como usar (dia a dia)
1. Dê duplo clique em **`INICIAR.bat`** → o aplicativo abre no navegador (deixe a janela preta aberta enquanto usa).
2. Na tela **Importar arquivos**, arraste os relatórios exportados do ERP — **qualquer nome de arquivo**, o tipo é reconhecido pelo conteúdo:
   - **1174** (vendas) → basta exportar os últimos dias. O período do arquivo novo substitui o mesmo período na base (reenviar não duplica).
   - **1211** (estoque) → sempre o completo, do dia.
   - **Vendas perdidas** → últimos dias ou período maior, tanto faz.
3. Clique em **Processar**. Em ~30 s (ou ~1,5 min com a planilha Excel) todas as telas se atualizam.

> Alternativa sem navegador: colocar os arquivos na pasta `entrada\` e rodar `ATUALIZAR.bat`.

## Consulta de Peças (`Consulta Vendas.exe`)
Programa de janela, 100% local, para consultar rapidamente uma peça: estoque e últimas vendas **por filial**.
1. Deixe os relatórios na **mesma pasta do .exe**: `1174` (.xlsx), `1211` (.csv) e `Vendas Perdidas` (.xlsx). O tipo é reconhecido pelo nome (contendo 1174 / 1211 / perdid) ou pelo cabeçalho. Vários 1174 ou Vendas Perdidas são somados (o período do mais novo prevalece); do 1211 vale o mais recente.
2. Dê duplo clique em **`Consulta Vendas.exe`** (abre em ~10 s).
3. Cole o código (Ctrl+V) — a consulta sai na hora. Pode colar vários códigos (coluna do Excel) e navegar entre eles, ou digitar parte do código/descrição.
- Abas:
  - **📊 Visão geral** — gráfico de vendas mês a mês (24 meses + venda perdida + previsão de 4 meses com faixa provável) e gráfico por filial (vendido 12m × estoque × previsão). Passe o mouse para ver os números; clique numa filial para filtrar tudo por ela.
  - **Resumo por filial** — estoque, última venda e **último cliente**, custo médio 12m, vendido 30d/90d/12m, nº de clientes, previsão do próximo mês, vendas perdidas.
  - **👥 Clientes por filial** — cada filial com seus clientes (conta, nº de compras, qtd, valor, última compra, último preço, vendedor, cidade) + gráfico dos maiores compradores.
  - **📈 Previsão** — previsão mês a mês por filial (do modelo do projeto, `base\modelo.pkl`), faixa provável, cobertura do estoque, data prevista de falta, situação e sugestão de compra. Para atualizar a previsão rode o `ATUALIZAR.bat` / aplicativo de previsão. Peça fora do modelo recebe uma estimativa simples (média 3 meses + mesmo mês do ano anterior).
  - **Últimas vendas** (com cliente, vendedor e cidade) e **Vendas perdidas**.
- **Custo médio vendido** (coluna *Custo Médio Venda* do 1174, que vem como total da linha): custo ÷ quantidade das vendas dos últimos 12 meses (sem venda no período, usa todo o 1174; sem nenhuma venda, o custo do estoque do 1211). Aparece no indicador do topo (com a margem), no Resumo por filial, por linha em Últimas vendas e na Previsão como **Valor (custo médio)** = quantidade a comprar × custo médio.
  - Ao trocar de versão o 1174 é relido uma vez (~3 a 4 minutos) para pegar a coluna de custo.
- **Seletor da previsão** (barra acima das abas, vale para todas): `Previsão por: [Período do 1174 | Modelo do projeto]   Base: [mês] até [mês]   Atender: [N meses]`.
  - *Período do 1174* (padrão: últimos 12 meses, atender 3): previsão de cada filial = média mensal (venda líquida + venda perdida) no período-base × meses a atender; **Comprar = previsão − estoque atual**. O período-base aparece destacado no gráfico. Escolher um mês já passa a previsão para este modo.
  - *Modelo do projeto*: previsão do `base\modelo.pkl` (até 4 meses), com situação e sugestão de compra do modelo.
- **Transferências** (botão no topo): sugestões de toda a rede, peça a peça, para o período-base e o *Atender* escolhidos (os mesmos da consulta).
  - A loja de **origem** guarda a própria previsão de *Atender + 1 mês*; só o que passar disso pode ir para outra loja. Peça parada (sem venda no período) sai primeiro.
  - A loja de **destino** recebe até cobrir a previsão de *Atender* meses. Previsão = média mensal (venda líquida + venda perdida) do período-base, a mesma da aba Previsão.
  - Filtros: Subgrupo (digite parte do nome e dê Enter), Origem, Destino, valor mínimo, busca. **Recorrência** = só sugere para a loja que vendeu ou perdeu venda em pelo menos N meses do período (padrão 2), para não mandar peça que vendeu uma única vez.
  - Gráfico por loja: clique = o que a loja envia; botão direito = o que recebe. Clique numa linha mostra o motivo; duplo clique abre a peça. "Exportar Excel" leva todas as linhas (a tela mostra as 3.000 de maior valor).
- **Saúde do estoque** (botão no topo): situação de cada loja, com o mesmo período, Atender, Subgrupo e Recorrência.
  - Por loja: valor em estoque (custo do 1211), itens, cobertura (estoque ÷ demanda mensal), giro 12m, **parado**, **excesso** (só a parte acima do limite, padrão 6 meses, ajustável), itens abaixo, **rupturas**, venda perdida e valor transferível.
  - Situação da peça: *Ruptura* (sem estoque, demanda recorrente) · *Abaixo* (estoque < previsão de Atender) · *Em giro* · *Excesso* · *Parada* (estoque sem nenhuma venda nem venda perdida no período).
  - Clique numa loja (tabela ou gráfico) para ver as peças dela; os botões de situação filtram a lista.
  - Venda perdida absurda (acima de 10× a maior venda da peça numa linha do 1174, mínimo 50 un) é ignorada nessas duas telas — ex.: um registro de 825.270 un de correia na filial 11.
- Linha rosa = a filial vende a peça mas está sem estoque. "Copiar aba (Excel)" copia a tabela da aba aberta.
- Trocou um relatório? Clique em **Recarregar arquivos**. A 1ª leitura do 1174 completo leva ~4 min; depois fica em cache (`base\cache_consulta\`) e abre em segundos.
- Para gerar o .exe de novo após mudar o código: `sistema\gerar_exe_consulta.bat`.

## Telas do aplicativo
| Tela | Para que serve |
|---|---|
| **Painel** | Alertas por subgrupo e filial, mapa subgrupo × filial, lista de itens (clique numa linha para abrir o item) |
| **Item a item** | Busca por código/descrição: vendas mensais × previsão (com faixa de erro), o que o modelo teria previsto no passado, projeção do estoque dia a dia (com e sem a compra sugerida) e o produto em todas as filiais |
| **Acurácia do modelo** | Acurácia medida por backtest em 3 níveis (item × filial, item na rede, subgrupo × filial), por mês, subgrupo e filial; comparativo de métodos; acompanhamento real mês a mês |
| **Importar arquivos** | Arrastar relatórios, processar, histórico de importações e downloads (Excel, painel offline, códigos substituídos) |

## O modelo de previsão (`sistema\modelo.py`)
- Previsão **mensal por filial × produto**, 4 meses à frente.
- Demanda = venda líquida + garantia + **vendas perdidas**. Picos isolados são cortados.
- **Códigos substituídos** (ex.: `81331R61BR` → `81331R61BRDSH`) têm o histórico unido automaticamente. Para incluir/bloquear ligações: `sistema\equivalencias_manual.csv`.
- 20 métodos competem num **backtest dos últimos 9 meses**; para cada perfil de demanda (SUAVE, ERRÁTICO, INTERMITENTE, IRREGULAR, RARO, INATIVO) fica o de menor erro com viés até ±30%.
- A cada mês novo, a primeira previsão fica guardada em `base\previsoes_historico\` e é comparada com a venda real quando o mês fecha (tela Acurácia → Acompanhamento real).

## Parâmetros
- Negócio (lead time, segurança, ciclo de compra): `sistema\config.py`
- Modelo (janelas, corte de picos, limite de viés): topo de `sistema\modelo.py`

## Estrutura
- `base\` — base consolidada e resultados do modelo (não apagar). `base\estoque_historico\` guarda uma foto do estoque por dia.
- `entrada\` — arquivos aguardando processamento; `processados\` guarda os já importados.
- `saida\` — planilha Excel, painel offline (HTML para enviar por e-mail) e lista de códigos substituídos.
- `sistema\` — código.
