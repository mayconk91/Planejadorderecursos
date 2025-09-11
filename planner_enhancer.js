/* planner_enhancer.js — improved collapsible behaviour for Planejamento
 * This script adds collapsible capabilities to the Planning tab without hiding
 * critical panels such as the resources/activities grid and the aggregated capacity charts.
 * Panels containing tables (#tblRecursos, #tblAtividades) or aggregated charts (#aggPanel or #aggCharts)
 * remain visible at all times. Other panels can be collapsed (except the first two) to reduce
 * visual clutter. When a panel with charts is expanded, chart dimensions are refreshed.
 */
(function(){
  /** Refresh capacity charts if global functions exist. */
  function refreshCharts() {
    // Attempt to refresh aggregated capacity charts and KPIs
    try { if (typeof window.renderCapacidadeAgregada === 'function') window.renderCapacidadeAgregada(); } catch(e){}
    try { if (typeof window.renderKPIs === 'function') window.renderKPIs(); } catch(e){}
    // Trigger a generic resize event which some chart libraries listen to
    try { window.dispatchEvent(new Event('resize')); } catch(e){}
  }

  document.addEventListener('DOMContentLoaded', () => {
    const tabPlan = document.getElementById('tab-plan');
    if (!tabPlan) return;
    // Gather direct child panels under the planning tab
    const sections = Array.from(tabPlan.children).filter(el => el.tagName && el.tagName.toLowerCase() === 'section' && el.classList.contains('panel'));
    sections.forEach((sec, idx) => {
      // Only proceed if the panel has a heading element
      const hdr = sec.querySelector('h2, h3, h4');
      if (!hdr) return;
      // Determine if this panel should not be collapsed:
      // - Contains resource or activity tables (edit controls live here)
      // - Is the aggregated capacity panel or contains aggregated charts
      const skip = sec.id === 'aggPanel' || sec.querySelector('#aggCharts') || sec.querySelector('#tblRecursos') || sec.querySelector('#tblAtividades');
      // Apply the collapsible class for styling
      sec.classList.add('collapsible');
      if (skip) {
        // Ensure skip panels remain expanded and cannot be collapsed
        sec.classList.remove('collapsed');
        hdr.style.cursor = 'default';
        return;
      }
      // For other panels, collapse all except the first two by default
      if (idx >= 2) {
        sec.classList.add('collapsed');
      }
      // Toggle collapsed state on header click
      hdr.style.cursor = 'pointer';
      hdr.addEventListener('click', () => {
        sec.classList.toggle('collapsed');
        // If panel just expanded, refresh charts to recalc sizes
        if (!sec.classList.contains('collapsed')) {
          refreshCharts();
        }
      });
    });
    // Initial refresh of charts after applying collapsible modifications
    refreshCharts();
  });
})();