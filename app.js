// app.js
// Excel dosyasından isim arayıp ilgili satırdaki günü listeleyen temel JS kodu

document.getElementById('searchBtn').addEventListener('click', function () {
  const fileInput = document.getElementById('excelFile');
  const rawName = document.getElementById('nameInput').value.trim();
  // Sadece harf ve Türkçe karakter kontrolü (büyük/küçük harf duyarsız)
  const namePattern = /^[a-zA-ZçÇğĞıİöÖşŞüÜ\s]+$/;
  if (!namePattern.test(rawName) || rawName.length === 0) {
    alert('Lütfen geçerli bir isim girin (sadece harf ve boşluk).');
    return;
  }
  // Arama ve karşılaştırma için inputu küçük harfe çeviriyoruz
  const name = rawName.toLocaleLowerCase('tr-TR');
  const resultList = document.getElementById('resultList');
  resultList.innerHTML = '';

  if (!fileInput.files[0] || !name) {
    alert('Lütfen dosya seçin ve isim girin.');
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (rows.length === 0) {
      alert('Excel dosyası boş.');
      return;
    }
    const isHasH2 = document.querySelector('h3');
    if (isHasH2) {
      document.body.removeChild(isHasH2);
    }
    let found = false;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (row.some((cell) => cell && cell.toString().toLocaleLowerCase('tr-TR') === name)) {
        found = true;
        const li = document.createElement('li');
        let displayRow = [...row];
        let dateStr = '';
        let dayName = '';
        // İlk hücre sayıysa, Excel tarihini dönüştür
        if (typeof displayRow[0] === 'number') {
          const excelDate = displayRow[0];
          const jsDate = new Date((excelDate - 25569) * 86400 * 1000);
          const day = String(jsDate.getDate()).padStart(2, '0');
          const month = String(jsDate.getMonth() + 1).padStart(2, '0');
          const year = jsDate.getFullYear();
          dateStr = `${day}.${month}.${year}`;
          // Haftanın gününü bul
          const daysTr = [
            'Pazar',
            'Pazartesi',
            'Salı',
            'Çarşamba',
            'Perşembe',
            'Cuma',
            'Cumartesi',
          ];
          dayName = daysTr[jsDate.getDay()];
          displayRow[0] = `<strong style='color:#1976d2'>${dateStr} (${dayName})</strong> <br/>`;
        } else if (typeof displayRow[0] === 'string') {
          // Eğer tarih string olarak gelirse, olduğu gibi göster
          displayRow[0] = `<strong>${displayRow[0]}</strong> <br/>`;
        }
        // Fazladan | işaretlerini tek bir tane olacak şekilde düzelt
        let rowHtml = displayRow.join(' | ').replace(/(\|\s*){2,}/g, ' | ');
        li.innerHTML = rowHtml;
        resultList.appendChild(li);
      }
    }
    if (!found) {
      const h2 = document.createElement('h3');
      h2.textContent = 'İsim bulunamadı.';
      document.body.appendChild(h2);
    }
  };
  reader.readAsArrayBuffer(fileInput.files[0]);
});
