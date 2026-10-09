# PPR Comex · agente de metas

Agente para entregar as metas mensais do PPR de comércio exterior. O gerente manda as planilhas
que já usa; o agente cruza tudo, aponta erros, monta a planilha-mestre e escreve o relatório.

## Arquivos

| Arquivo | Para quê |
|---|---|
| `AGENTE.md` | Instruções do agente: metas, o que pedir, regras de cálculo, formato das entregas |
| `../.claude/agents/ppr-comex.md` | O agente registrado no Claude Code |
| `scripts/ler_arquivos.py` | Leitores das planilhas originais (periodicidade, S.A., fluxo, gastos, câmbio, uso e estoque de embalagens) |
| `scripts/fechar_mes.py` | Fechamento do mês: planilha-mestre + relatório + pendências |
| `scripts/documento.py` | Saída em Word/Markdown e formatação da planilha |
| `dados/` (fora do git) | `origem/` com as planilhas do gerente, `PPR_Comex_<ano>.xlsx` (mestre) e `apelidos.csv` |
| `relatorios/` (fora do git) | Um diretório por mês com o relatório, o resumo para o gestor e as pendências |

## Como usar

**No Claude Code:** "use o agente ppr-comex para fechar outubro" e anexe as planilhas do mês.

**Manualmente:**

```
python3 ppr-comex/scripts/fechar_mes.py --mes 2026-10
```

Preencha as células amarelas da planilha-mestre (pedidos do ano, motivos, conclusão e cobranças
das S.A., ações de custo) e rode de novo. O que foi digitado é preservado.
