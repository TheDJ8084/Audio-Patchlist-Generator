function onOpen() {
  DocumentApp.getUi()
      .createMenu('Patchlist Generator')
      .addItem('Configure patchlist', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Patchlist Generator')
      .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

function createTables(data) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Optioneel: uncomment deze regel om bestaande content te wissen
  // body.clear();
  Logger.log('--- Start createTables functie ---');
  Logger.log('Ontvangen data: ' + JSON.stringify(data));

  var consoleName = data.consoleName.trim();
  if (consoleName) {
    consoleName = consoleName + ' ';
  } else {
    consoleName = '';
  }

  var types = ['Console', 'StageBox', 'Aux'];
  var anyTableCreated = false;

  // --- Sectie voor INPUT tabellen ---
  var hasInputsToCreate = types.some(type => {
    var prefix = type.charAt(0).toLowerCase() + type.slice(1);
    return data[prefix + 'Toggle'] && parseInt(data[prefix + 'InputRows'] || 0) > 0;
  });

  if (hasInputsToCreate) {
    if (anyTableCreated) {
      body.appendPageBreak();
      Logger.log('Paginabrake toegevoegd voor start Inputs sectie.');
    }
    
    body.appendParagraph(consoleName + 'Inputs').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    anyTableCreated = true;
    Logger.log('Algemene kop: ' + consoleName + 'Inputs geplaatst.');

    var firstInputTableInSector = true;
    types.forEach(function(type) {
      var prefix = type.charAt(0).toLowerCase() + type.slice(1);
      var toggleProp = prefix + 'Toggle';
      var inputRowsProp = prefix + 'InputRows';

      var isToggleActive = data[toggleProp];
      var inputRows = parseInt(data[inputRowsProp] || 0);

      if (isToggleActive && inputRows > 0) {
        if (!firstInputTableInSector) {
          body.appendPageBreak();
          Logger.log('  Paginabrake toegevoegd tussen input-tabellen voor ' + type + ' Inputs.');
        }

        Logger.log('  Aanmaken ' + type + ' Inputs tabel met ' + inputRows + ' rijen.');
        body.appendParagraph(type + ' Inputs').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        
        // --- AANPASSING HIER ---
        if (type === 'Aux') {
          createTable(body, inputRows, ['CH', 'Source', 'Description']); // 48V verwijderd
        } else {
          createTable(body, inputRows, ['CH', 'Source', '48V', 'Description']);
        }
        
        firstInputTableInSector = false;
      } else {
        Logger.log('  Geen ' + type + ' Inputs tabel aangemaakt (toggle inactief of rijen <= 0).');
      }
    });
  } else {
    Logger.log('Geen INPUT tabellen om aan te maken.');
  }

  // --- Sectie voor OUTPUT tabellen ---
  var hasOutputsToCreate = types.some(type => {
    var prefix = type.charAt(0).toLowerCase() + type.slice(1);
    return data[prefix + 'Toggle'] && parseInt(data[prefix + 'OutputRows'] || 0) > 0;
  });

  if (hasOutputsToCreate) {
    if (anyTableCreated) {
      body.appendPageBreak();
      Logger.log('Paginabrake toegevoegd voor start Outputs sectie.');
    }

    body.appendParagraph(consoleName + 'Outputs').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    anyTableCreated = true;
    Logger.log('Algemene kop: ' + consoleName + 'Outputs geplaatst.');

    var firstOutputTableInSector = true;
    types.forEach(function(type) {
      var prefix = type.charAt(0).toLowerCase() + type.slice(1);
      var toggleProp = prefix + 'Toggle';
      var outputRowsProp = prefix + 'OutputRows';

      var isToggleActive = data[toggleProp];
      var outputRows = parseInt(data[outputRowsProp] || 0);

      if (isToggleActive && outputRows > 0) {
        if (!firstOutputTableInSector) {
          body.appendPageBreak();
          Logger.log('  Paginabrake toegevoegd tussen output-tabellen voor ' + type + ' Outputs.');
        }
        
        Logger.log('  Aanmaken ' + type + ' Outputs tabel met ' + outputRows + ' rijen.');
        body.appendParagraph(type + ' Outputs').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        createTable(body, outputRows, ['CH', 'Destination', 'Description']);
        firstOutputTableInSector = false;
      } else {
        Logger.log('  Geen ' + type + ' Outputs tabel aangemaakt (toggle inactief of rijen <= 0).');
      }
    });
  } else {
    Logger.log('Geen OUTPUT tabellen om aan te maken.');
  }

  DocumentApp.getUi().alert('Patchlist has been succesfully created');
  Logger.log('--- Einde createTables functie ---');
}

function createTable(body, numRows, headers) {
  Logger.log('  Start createTable: numRows=' + numRows + ', headers=' + headers.join(', '));
  
  if (numRows <= 0) {
    Logger.log('  Geen rijen om aan te maken in tabel.');
    return;
  }

  var table = body.appendTable();
  var headerRow = table.appendTableRow();

  headers.forEach(function(headerText) {
    headerRow.appendTableCell(headerText).setBold(true);
  });

  for (var i = 0; i < numRows; i++) {
    var row = table.appendTableRow();
    for (var j = 0; j < headers.length; j++) {
      if (j === 0) { // Eerste kolom
        // Aanpassing hier: converteer het getal naar een string
        row.appendTableCell(String(i + 1)); 
      } else { // Overige kolommen
        row.appendTableCell('');
      }
    }
  }
  Logger.log('  Tabel succesvol aangemaakt.');
}
