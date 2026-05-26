# KPI Comercial - Dashboard

Dashboard comercial com foco em:
- Painel de Vendas (gráficos e filtros)
- CRM (funil cumulativo, kanban e tarefas)
- Meta (velocímetro de atingimento + comparativo em R$)

## Estrutura funcional

- `Base_Painel_KPI_.xlsx`: base principal (fonte de dados)
- `scripts/extract_kpi_data.py`: extrator Excel -> JSON
- `scripts/dashboard_server.py`: servidor local + endpoint `POST /api/refresh`
- `dashboard/`: aplicação web

## Atualizar base e abrir localmente

Use o atalho:

```bat
hostlocal.bat
```

Ele executa:
1. Atualização de `dashboard/data/kpi_data.json`
2. Inicialização do servidor local
3. Abertura automática do dashboard no navegador

## Execução manual (opcional)

```powershell
python .\scripts\extract_kpi_data.py --input ".\Base_Painel_KPI_.xlsx" --output ".\dashboard\data\kpi_data.json"
python .\scripts\dashboard_server.py --port 5500
```

Acesse: `http://localhost:5500`

## Observações

- O dashboard não altera o arquivo Excel original.
- A base usada no front-end é sempre o JSON gerado.
- O endpoint `/api/refresh` regenera o JSON com a planilha atual.

## Publicação online (GitHub Pages)

O projeto já está preparado com workflow de deploy automático:

- Arquivo: `.github/workflows/deploy-pages.yml`
- Origem publicada: pasta `dashboard/`
- Deploy acionado em push para `main` e `principal` (ou manual via Actions)
- Antes de publicar, o workflow regenera `dashboard/data/kpi_data.json` a partir de `Base_Painel_KPI_.xlsx`

### Passos no GitHub (uma vez)

1. Em `Settings > Pages`, deixe a opção de fonte como `GitHub Actions`.
2. Faça push das alterações para `main`.
3. Acompanhe em `Actions` o workflow `Deploy Dashboard to GitHub Pages`.
4. Após concluir, acesse a URL publicada em `Settings > Pages` ou no log do job.

### Importante no ambiente online

- O botão **Atualizar Tudo** fica desabilitado fora de `localhost` porque o endpoint local `/api/refresh` não existe no Pages.
- Para atualizar dados online:
  1. gere `dashboard/data/kpi_data.json` localmente com a planilha;
  2. faça commit/push do JSON atualizado;
  3. o Pages publica automaticamente.
