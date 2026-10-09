"""Estrutura única da planilha PPR Comex.

Os nomes de abas e colunas ficam aqui para que o gerador da planilha
(criar_planilha.py) e o gerador de relatórios (relatorio_mensal.py) usem
sempre os mesmos cabeçalhos.
"""

# Colunas de entrada (o usuário preenche). Colunas calculadas ficam à direita
# e são definidas no criar_planilha.py como fórmulas.
CLIENTES = [
    "Cód. Cliente", "Cliente", "País", "Idioma", "Canal/Representante",
    "Periodicidade (meses)", "Status cadastro", "Observação",
]
PEDIDOS = [
    "Nº Pedido", "Data do pedido", "Cód. Cliente", "Produto", "Toneladas",
    "Valor (USD)", "Incoterm", "Embarque negociado", "Embarque realizado",
]
ACOMPANHAMENTO = [
    "Mês ref.", "Cód. Cliente", "Cliente", "Compra prevista em", "Situação",
    "Motivo", "Detalhe do motivo", "Ação combinada", "Responsável",
    "Data de retorno", "Status da ação", "Chance de recompra",
]
ATIVIDADES = [
    "Nº Pedido", "Cliente", "Departamento", "Atividade", "Data solicitação",
    "Prazo (dias)", "Data conclusão", "Nº cobranças", "Última cobrança",
    "Causa do atraso", "Observação",
]
CUSTOS = [
    "Mês ref.", "Categoria", "Descrição da ação", "Tipo de ganho",
    "Custo de referência (R$)", "Custo realizado (R$)", "Evidência", "Responsável",
]
EMBALAGENS = [
    "Código", "Descrição", "Tipo", "Produto/linha", "Idioma", "Mercado original",
    "Medidas (cm)", "Qtd em estoque", "Unidade", "Custo unit. (R$)",
    "Data de entrada", "Status", "Localização", "Observação",
]
USO_EMBALAGENS = [
    "Data", "Nº Pedido", "Cliente", "Código embalagem", "Qtd utilizada",
    "Forma de uso", "Aprovado pela produção", "Observação",
]

LISTAS = {
    "Status cadastro": ["Ativo", "Inativo", "Perdido", "Prospect"],
    "Periodicidade": [1, 2, 3, 4, 6, 12],
    "Situação": ["Não recomprou", "Atrasado", "Recomprou", "Novo/reativado",
                 "Antecipou", "Perdido"],
    "Motivo": [
        "Estoque alto no cliente", "Preço / câmbio", "Perdeu para concorrente",
        "Reclamação de qualidade", "Atraso em embarque anterior",
        "Barreira sanitária / regulatória", "Frete / contêiner / logística",
        "Demanda sazonal", "Crédito / financeiro do cliente",
        "Troca de comprador / estrutura", "Produto indisponível",
        "Pedido adiado para o mês seguinte", "Cliente encerrou atividade",
        "Sem retorno do cliente", "Outro",
    ],
    "Status da ação": ["Aberta", "Concluída", "Cancelada"],
    "Chance de recompra": ["Alta", "Média", "Baixa"],
    "Departamento": [
        "Comercial/Comex", "Compras/Suprimentos", "Marketing/Arte",
        "Qualidade/Regulatório", "PCP/Produção", "Logística/Embarque",
        "Financeiro", "Documentação/Despachante",
    ],
    "Categoria custo": [
        "Amostras", "Frete de amostras", "Inspeções", "Viagens internacionais",
        "Impressões / documentos", "Courier", "Frete internacional",
        "Embalagens reaproveitadas", "Certificados / taxas", "Outros",
    ],
    "Tipo de ganho": [
        "Economia direta", "Custo evitado", "Custo transferido ao cliente",
        "Receita recuperada",
    ],
    "Tipo embalagem": [
        "Caixa master", "Cartucho", "Saco/filme", "Etiqueta", "Bandeja",
        "Divisória/reforço", "Outro",
    ],
    "Status embalagem": ["Obsoleta", "Ativa", "Em avaliação"],
    "Forma de uso": [
        "Mesmo idioma", "Bonificação", "Reforço/acessório",
        "Reetiquetagem/neutra", "Descarte/reciclagem",
    ],
    "Sim/Não": ["Sim", "Não"],
}

# Prazo padrão de cada departamento para entregar a S.A. (dias corridos).
SLA_PADRAO = {
    "Comercial/Comex": 2, "Compras/Suprimentos": 10, "Marketing/Arte": 7,
    "Qualidade/Regulatório": 7, "PCP/Produção": 15, "Logística/Embarque": 5,
    "Financeiro": 3, "Documentação/Despachante": 5,
}

PARAMETROS = [
    ("Mês de referência", "2026-10-01", "Primeiro dia do mês que está sendo apurado"),
    ("Ano base", 2025, "Ano usado como base da positivação"),
    ("Ano meta", 2026, "Ano do PPR"),
    ("Meta de aumento de positivação", 0.20, "20% sobre o ano base"),
    ("Meta de volume (t/mês)", 75, "Média mensal"),
    ("Prazo de embarque padrão (dias)", 30, "Prazo negociado com o cliente"),
    ("Idade para obsoleta (meses)", 12, "Embalagem parada há mais tempo que isso"),
]
