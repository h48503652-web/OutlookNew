(function(){
  const form = document.getElementById('draftForm');
  const statusEl = document.getElementById('status');

  function setStatus(msg, type='info'){
    statusEl.textContent = msg;
    statusEl.className = type === 'error' ? 'error' : (type === 'success' ? 'success' : 'muted');
  }

  function splitRecipients(raw){
    return raw
      .split(/[;,\n]/)
      .map(s => s.trim())
      .filter(Boolean);
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    setStatus('שולח בקשה לשירות המקומי...');

    const subject = document.getElementById('subject').value.trim();
    const recipientsRaw = document.getElementById('recipients').value.trim();
    const body = document.getElementById('body').value;
    const fileInput = document.getElementById('attachment');
    const file = fileInput.files[0];

    if(!subject || !recipientsRaw || !body || !file){
      setStatus('יש למלא את כל השדות ולצרף קובץ.', 'error');
      return;
    }

    const recipients = splitRecipients(recipientsRaw);
    if(recipients.length === 0){
      setStatus('לא נמצאו נמענים תקינים.', 'error');
      return;
    }

    const fd = new FormData();
    fd.append('subject', subject);
    fd.append('body', body);
    fd.append('recipients', JSON.stringify(recipients));
    fd.append('attachment', file);

    try{
      const res = await fetch('http://127.0.0.1:5005/draft', {
        method: 'POST',
        body: fd
      });
      const data = await res.json().catch(()=>({}));
      if(!res.ok){
        throw new Error(data.message || 'שגיאה ביצירת טיוטות');
      }
      setStatus(`נפתחו ${data.created || recipients.length} טיוטות ב-Outlook.`, 'success');
    }catch(err){
      setStatus(`שגיאה: ${err.message}. ודאי שהשירות המקומי רץ.`, 'error');
    }
  });
})();
