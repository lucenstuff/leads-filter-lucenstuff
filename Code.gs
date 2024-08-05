function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function filterLeads(form) {
  var sheet = SpreadsheetApp.openById('1Za57-U3Sj_YfZsMe5BTYa5WLFKIil4fvU9XSMmuBKew').getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var downloadCountIndex = headers.indexOf('DownloadCount');

  if (downloadCountIndex == -1) {
    headers.push('DownloadCount');
    downloadCountIndex = headers.length - 1;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    for (var i = 1; i < data.length; i++) {
      data[i].push(0);
    }
    sheet.getRange(2, 1, data.length - 1, headers.length).setValues(data.slice(1));
  }

  var filteredData = data.filter((row, index) => {
    if (index === 0) return true; 

    var roleMatch = !form.role || (row[headers.indexOf('Role')] && row[headers.indexOf('Role')].toString().toLowerCase().includes(form.role.toLowerCase()));
    var industryMatch = !form.industry || (row[headers.indexOf('Industry')] && row[headers.indexOf('Industry')].toString().toLowerCase().includes(form.industry.toLowerCase()));
    var countryMatch = !form.country || (row[headers.indexOf('Country')] && row[headers.indexOf('Country')].toString().toLowerCase().includes(form.country.toLowerCase()));
    var cnaeMatch = !form.cnae || (row[headers.indexOf('CNAE')] && row[headers.indexOf('CNAE')].toString().toLowerCase().includes(form.cnae.toLowerCase()));
    
    return roleMatch && industryMatch && countryMatch && cnaeMatch;
  });


  filteredData.forEach((row, index) => {
    if (index !== 0) {
      row[downloadCountIndex]++;
    }
  });

  var allData = [headers].concat(data.slice(1));
  sheet.getRange(1, 1, allData.length, allData[0].length).setValues(allData);


  var csvContent = filteredData.map(row => row.join(',')).join('\n');
  var blob = Utilities.newBlob(csvContent, 'text/csv', 'filtered_leads.csv');
  var file = DriveApp.createFile(blob);

  return file.getUrl();
}
