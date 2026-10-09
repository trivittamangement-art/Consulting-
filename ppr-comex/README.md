# PPR Comex · agente de metas

Agente para entregar as metas mensais do PPR de comércio exterior: você manda as informações
básicas e ele monta a planilha, os relatórios e os textos.

## O que tem aqui

| Arquivo | Para quê |
|---|---|
| `AGENTE.md` | Instruções do agente: o que ele pede, quando, como calcula e o que entrega |
| `../.claude/agents/ppr-comex.md` | O mesmo agente registrado no Claude Code |
| `modelos/PPR_Comex_Controle.xlsx` | Planilha-mestre em branco (Painel, Clientes, Pedidos, Acompanhamento, S_Atividades, Custos, Embalagens) |
| `modelos/PPR_Comex_EXEMPLO.xlsx` | A mesma planilha com dados **fictícios**, para ver funcionando |
| `modelos/exemplo_relatorio_2026-09/` | Relatório, descritivo de embalagens e pendências gerados a partir do exemplo |
| `scripts/relatorio_mensal.py` | Gera o relatório do mês a partir da planilha |
| `scripts/criar_planilha.py` | Recria a planilha em branco |

## Como usar

**No Claude Code (este repositório):** escreva algo como "use o agente ppr-comex para fechar
setembro". Na primeira vez ele pede a sua planilha de 2025, os pedidos de 2026, um exemplo de
S.A., o inventário de embalagens e alguns custos de referência. Seus dados ficam em
`ppr-comex/dados/` e os relatórios em `ppr-comex/relatorios/`. As duas pastas estão no
`.gitignore` e não vão para o GitHub.

**No claude.ai (Projeto):** crie um Projeto "PPR Comex", cole o conteúdo de `AGENTE.md` nas
instruções do Projeto e suba a `PPR_Comex_Controle.xlsx` (ou a sua planilha atual) nos arquivos.
Todo mês, abra uma conversa nova no Projeto e escreva "fechar <mês>".

## O que você manda todo mês (kit do mês)

1. Pedidos lançados no mês (exportação do sistema)
2. Embarques realizados (pedido + data)
3. S.A. do mês (exportação ou print) e as cobranças que você fez
4. O que aconteceu com amostras, inspeções, viagens, impressões e fretes
5. Embalagens que a produção usou (e o inventário, se mudou)

O agente devolve o relatório das 3 metas em Word, o descritivo de embalagens para a produção,
a planilha atualizada, uma lista curta de motivos de não recompra para você responder em uma
linha por cliente e um resumo de 5 linhas para o seu gestor. No dia 15 ele adianta os clientes
que vencem no mês seguinte (com mensagem pronta no idioma de cada um) e os e-mails de cobrança
das S.A. atrasadas.
