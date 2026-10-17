let people = [];

function downloadTemplate() {
  const ws = XLSX.utils.aoa_to_sheet([
    ['id', 'adsoyad', 'telefon_no'],
    [1, 'Ahmet Yılmaz', '905321234567'],
    [2, 'Ayşe Kaya', '905331234567'],
  ]);
  ws['!cols'] = [{ wch: 6 }, { wch: 20 }, { wch: 18 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Kişiler');
  XLSX.writeFile(wb, 'whatsapp_davet_sablonu.xlsx');
}

function loadExcel() {
  const file = document.getElementById('excelFile').files[0];
  if (!file) return alert('Lütfen bir Excel dosyası seçin.');

  const reader = new FileReader();
  reader.onload = function (e) {
    const workbook = XLSX.read(e.target.result, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    if (!rows.length || !('telefon_no' in rows[0])) {
      return alert('Excel dosyasında "telefon_no" sütunu bulunamadı. Lütfen sütun başlıklarını kontrol edin.');
    }

    people = rows.map((r, i) => ({
      id: r.id ?? i + 1,
      adsoyad: r.adsoyad ?? '',
      telefon_no: String(r.telefon_no).replace(/\D/g, ''),
      sent: false
    }));

    renderTable();
  };
  reader.readAsArrayBuffer(file);
}

function renderTable() {
  const tbody = document.getElementById('personTable');
  tbody.innerHTML = '';
  people.forEach((p, i) => {
    const tr = document.createElement('tr');
    tr.id = `row-${i}`;
    tr.innerHTML = `
      <td>${p.id}</td>
      <td>${p.adsoyad}</td>
      <td>${p.telefon_no}</td>
      <td><span class="badge ${p.sent ? 'badge-sent' : 'badge-wait'}" id="status-${i}">${p.sent ? '✅ Gönderildi' : '⏳ Bekliyor'}</span></td>
      <td><button class="btn btn-primary" onclick="sendOne(${i})">Gönder</button></td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('summary').textContent = `Toplam ${people.length} kişi yüklendi.`;
  document.getElementById('tableArea').style.display = 'block';
}

function buildWaUrl(telefon_no) {
  const groupLink = document.getElementById('groupLink').value.trim();
  const extraMsg  = document.getElementById('extraMsg').value.trim();
  if (!groupLink) { alert('Lütfen grup davet linkini girin.'); return null; }

  const msg = encodeURIComponent(`${extraMsg ? extraMsg + '\n' : ''}${groupLink}`);
  return `https://web.whatsapp.com/send?phone=${telefon_no}&text=${msg}`;
}

function sendOne(i) {
  const url = buildWaUrl(people[i].telefon_no);
  if (!url) return;
  window.open(url, '_blank');
  people[i].sent = true;
  const badge = document.getElementById(`status-${i}`);
  badge.textContent = '✅ Gönderildi';
  badge.className = 'badge badge-sent';
}

function sendAll() {
  const groupLink = document.getElementById('groupLink').value.trim();
  if (!groupLink) return alert('Lütfen grup davet linkini girin.');
  if (!people.length) return alert('Önce Excel dosyasını yükleyin.');

  const unsent = people.map((_, i) => i).filter(i => !people[i].sent);
  if (!unsent.length) return alert('Tüm kişilere zaten gönderildi.');

  let current = 0;

  function sendNext() {
    if (current >= unsent.length) return;
    sendOne(unsent[current]);
    current++;
    if (current < unsent.length) {
      window.addEventListener('focus', function onFocus() {
        window.removeEventListener('focus', onFocus);
        setTimeout(sendNext, 500);
      });
    }
  }

  sendNext();
}
