"""Parâmetros da previsão (BETA). Ajuste aqui e rode ATUALIZAR.bat de novo."""

# Prazo médio de reposição (dias entre pedir e a peça chegar na filial)
LEAD_TIME_DIAS = 30

# Estoque de segurança, em dias de demanda, somado ao lead time
SEGURANCA_DIAS = 15

# Ciclo de compra: quantos dias de demanda cada pedido deve cobrir além do LT+segurança
CICLO_COMPRA_DIAS = 30

# Janela "recente" usada para medir o ritmo de vendas atual
JANELA_RECENTE_DIAS = 90

# Demanda mínima esperada no horizonte (lead time + segurança) para o item entrar como
# RUPTURA/CRÍTICO/ATENÇÃO. Abaixo disso, a previsão é residual (item quase parado).
MIN_UNID_HORIZONTE = 0.5

# Peso da sazonalidade (mesmo período do ano anterior) na demanda prevista.
# 0 = só ritmo recente; 1 = só ano anterior. Agro é sazonal, então começa em 0.5.
PESO_SAZONAL = 0.5

# Mínimo de meses distintos com venda nos últimos 12 meses para o item ser
# considerado "recorrente". Itens abaixo disso aparecem como BAIXA RECORRÊNCIA.
MIN_MESES_RECORRENCIA = 3

# Venda perdida conta como demanda não atendida (1.0 = conta 100% da quantidade)
PESO_VENDA_PERDIDA = 1.0

# Tipos de lançamento do 1174 que formam a demanda (devolução já vem negativa e abate;
# "Outras saídas" = garantia, consumo real). "Outras entradas" (serviço) fica de fora.
TIPOS_DEMANDA = ["Vendas", "Devolução de vendas", "Outras saídas"]

# Códigos que não são peça de estoque (ficam fora da previsão)
EXCLUIR_PREFIXOS = ["LIC/"]

NOMES_FILIAIS = {
    1: "SM", 2: "SG", 3: "BG", 4: "AL", 5: "JC", 6: "PL", 7: "UR",
    8: "IT", 9: "SB", 10: "LJ", 11: "CS", 12: "SV", 13: "ES",
}
