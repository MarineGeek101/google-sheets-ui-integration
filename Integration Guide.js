function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Student Analytics Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getStudentData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targetName = 'Overall % upto 4th Sem';
    var sheet = null;

    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].getName().trim().toLowerCase() === targetName.trim().toLowerCase()) {
        sheet = sheets[s];
        break;
      }
    }

    if (!sheet) {
      return { error: 'Sheet tab "' + targetName + '" was not found in this spreadsheet.' };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];

    var numRows = lastRow - 2;
    if (numRows <= 0) return [];

    // Range A to AC (29 columns total)
    var values = sheet.getRange(3, 1, numRows, 29).getDisplayValues();
    var photoFormulas = sheet.getRange(3, 3, numRows, 1).getFormulas(); // Photo formula check in Column C

    var cleanData = [];

    function cellString(rowArr, idx) {
      if (!rowArr || idx >= rowArr.length || rowArr[idx] === null || rowArr[idx] === undefined) return '';
      return String(rowArr[idx]).trim();
    }

    function extractDriveId(url) {
      if (!url) return '';
      var match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || 
                  url.match(/[?&]id=([a-zA-Z0-9_-]+)/) || 
                  url.match(/open\?id=([a-zA-Z0-9_-]+)/) || 
                  url.match(/[-\w]{25,}/);
      return match ? (match[1] || match[0]) : '';
    }

    function makePhotoUrl(url) {
      if (!url) return '';
      url = String(url).trim();
      if (url.indexOf('drive.google.com') === -1 && url.indexOf('googleusercontent.com') === -1) {
        return url;
      }
      var fileId = extractDriveId(url);
      return fileId ? ('https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1000') : url;
    }

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var cellIndex = i + 3;

      var rawId = cellString(row, 1);   // Col B: Student ID
      var rawName = cellString(row, 3); // Col D: Name

      if (rawId === '' && rawName === '') continue;

      var rawPct = cellString(row, 23); // Col X: Percentage
      var pctStr = '0.00';
      if (rawPct !== '') {
        var parsedPct = parseFloat(rawPct.replace(/%/g, ''));
        pctStr = isNaN(parsedPct) ? '0.00' : parsedPct.toFixed(2);
      }

      var photoUrl = '';
      var rowFormula = photoFormulas[i] && photoFormulas[i][0] ? photoFormulas[i][0] : '';

      if (rowFormula && rowFormula.toUpperCase().indexOf('=IMAGE') === 0) {
        var match = rowFormula.match(/https?:\/\/[^"')\s,]+/i);
        if (match && match[0]) {
          photoUrl = match[0].replace(/["')]+$/, '');
        }
      }

      if (!photoUrl) {
        var rawPhoto = cellString(row, 2); // Col C: Photos
        if (rawPhoto && rawPhoto.indexOf('http') === 0) {
          photoUrl = rawPhoto;
        }
      }

      photoUrl = makePhotoUrl(photoUrl);

      cleanData.push({
        sn: cellString(row, 0) || String(cleanData.length + 1),
        id: rawId || 'N/A',
        photo: photoUrl,
        name: rawName || 'Row ' + cellIndex,
        division: cellString(row, 4) || '-',

        max1: cellString(row, 5) || '-', obt1: cellString(row, 6) || '-', pct1: cellString(row, 7) || '-', rk1: cellString(row, 8) || '-',
        max2: cellString(row, 9) || '-', obt2: cellString(row, 10) || '-', pct2: cellString(row, 11) || '-', rk2: cellString(row, 12) || '-',
        max3: cellString(row, 13) || '-', obt3: cellString(row, 14) || '-', pct3: cellString(row, 15) || '-', rk3: cellString(row, 16) || '-',
        max4: cellString(row, 17) || '-', obt4: cellString(row, 18) || '-', pct4: cellString(row, 19) || '-', rk4: cellString(row, 20) || '-',

        tMax: cellString(row, 21) || '-',
        tObt: cellString(row, 22) || '-',
        pct: pctStr,
        overRank: cellString(row, 24) || '-',

        sem4Credit: cellString(row, 25) || '-',
        sem4Discredit: cellString(row, 26) || '-',
        totalSem4CP: cellString(row, 27) || '-',
        sem4RankCP: cellString(row, 28) || '-'
      });
    }
    return cleanData;
  } catch (err) {
    return { error: 'Apps Script Trace Crash: ' + err.toString() };
  }
}