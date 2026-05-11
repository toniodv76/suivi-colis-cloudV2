async function loadColis(){
  const r = await fetch('/api/colis');
  const data = await r.json();
  const table = document.getElementById('table');
  if(!data.length){ table.innerHTML = '<div class="empty">AUCUN COLIS</div>'; return; }
  table.innerHTML = data.map(c => `
    <div class="screen-row ${c.statut === 'FAIT' ? 'done' : 'todo'}">
      <div class="client">${escapeHtml(c.client)}</div>
      <div class="adresse">${escapeHtml(c.adresse)}</div>
      <div class="contenu">${escapeHtml(c.contenu)}</div>
      <div class="statut">${escapeHtml(c.statut)}</div>
    </div>`).join('');
}
function escapeHtml(t){return (t||'').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));}
function clock(){document.getElementById('clock').textContent = new Date().toLocaleString('fr-FR');}
setInterval(loadColis, 2000); setInterval(clock, 1000); loadColis(); clock();
