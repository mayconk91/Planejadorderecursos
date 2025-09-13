PLANEJADOR DE RECURSOS — OFFLINE (v2)
=================================

O que vem no pacote:
- index.html        → Abrir no navegador (sem internet) para usar o sistema.
- app.js            → Código JS (vanilla) do planejador.
- styles.css        → Estilos básicos.
- (este) README.txt → Instruções rápidas.

Implementado:
1) Exportar Excel do estado atual (inclui IDs) + CSV para reimportação estável.
2) Filtros combinados (Status, Tipo, Senioridade) + busca por título.
3) Indicadores de capacidade por dia (heatmap) e empilhamento de atividades concorrentes.
4) CSV "Power BI" com 1 linha por dia por atividade.
5) Rodar 100% offline (sem dependências externas, dados no localStorage).
6) Sem botão de reset (excluir é item a item).
7) TRILHA de alterações de datas com justificativa obrigatória e "Usuário atual".
8) Exportar histórico por atividade e consolidado (todas as atividades).
9) Backup/Restaurar (JSON) do estado completo.
10) Validações: janela ativa por recurso (início/fim) e alerta ao ultrapassar 100%.
11) Capacidade agregada (semanal/mensal) com gráficos em canvas por recurso.

DICAS
-----
- Para definir janela ativa do recurso, use os campos "Início ativo" e "Fim ativo".
- Use o campo "Usuário atual" (canto superior direito) para registrar quem alterou.
- Para importar, use os CSVs gerados pelos próprios botões de export.

Ass.: Gerado em 2025-09-05


ABAS
----
- **Planejamento**: filtros, Gantt, tabelas e capacidade agregada.
- **Exportações & Backup**: Importar/Exportar CSV, Excel, Power BI, **Backup (JSON)** e **Restaurar** (importa o JSON de backup).


SALVAR EM PASTA (fora do navegador)
-----------------------------------
- No topo da página, clique em **📁 Definir/Alterar pasta de dados** e escolha a pasta onde o app está (ou outra).
- O sistema passa a **salvar automaticamente** `resources.json`, `activities.json` e `trails.json` nessa pasta a cada alteração. Existe também **💾 Salvar agora** e **🔄 Recarregar da pasta**.
- Requer navegador Chromium (Chrome/Edge) com suporte ao **File System Access API**. Em `file://` costuma funcionar; caso seu navegador impeça, abra via `http://localhost` (qualquer servidor local).
- Se não houver suporte/permissão, o app usa somente o armazenamento do navegador como fallback.


NOTA PDF (offline)
------------------
- O botão **Exportar PDF** funciona com jsPDF/html2canvas se presentes no navegador.
- Quando não disponíveis, usa **janela de impressão** (Ctrl+P) para permitir **Salvar como PDF** (100% offline).


## Multiusuário (contorno)
- Observa alterações nos arquivos da pasta compartilhada e recarrega automaticamente quando outra sessão salvar (polling ~3s).
- Recursos/Atividades/Trilhas e Gestão de Horas (dados_enhancer.json) são observados.
- Em conflitos simultâneos, prevalece o último salvamento.
