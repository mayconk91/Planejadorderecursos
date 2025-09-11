Correções aplicadas:
- Evita painel em branco na aba 'Gestão de Horas (Externos)' com detecção robusta de recursos vindos de window.resources, LocalStorage e tabela DOM (#tblRecursos).
- Observa a tabela de recursos para re-render automático quando dados chegarem, evitando tela vazia.
- Inputs usam <input type="time" step="60"> para horas, garantindo suporte a HH:MM e minutos.
- Ajuste de KPI: atividades 'Concluída' são ignoradas no cálculo de sobrecarga por intervalo.
- Toolbar de pasta (File System Access) preservada e funcional.
