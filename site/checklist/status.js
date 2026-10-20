/* When a document is due, judged against today. Shared by the checklist page and the team hub,
   so the two can never disagree about what is overdue. */
window.HCStatus = (function(){
  var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var LABEL = {overdue:'Overdue', soon:'Due within 30 days', none:'No review date on record', ok:'In date'};

  function iso(d){ return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
  function pretty(s){ if(!s) return ''; var p = s.split('-'); return (+p[2])+' '+MONTHS[+p[1]-1]+' '+p[0]; }

  var today = new Date(), soon = new Date(today);
  soon.setDate(soon.getDate()+30);
  var TODAY = iso(today), SOON = iso(soon);

  function of(d){
    if(d.fixed_status) return {key: d.fixed_status === 'Missing' ? 'missing' : 'fixed', label:d.fixed_status};
    if(!d.next_due) return {key:'none', label:LABEL.none};
    if(d.next_due < TODAY) return {key:'overdue', label:LABEL.overdue};
    if(d.next_due < SOON) return {key:'soon', label:LABEL.soon};
    return {key:'ok', label:LABEL.ok};
  }

  return {of: of, pretty: pretty, today: TODAY, LABEL: LABEL};
})();
