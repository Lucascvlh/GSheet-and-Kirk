function onOpen() {
 SpreadsheetApp
   .getUi()
   .createMenu("Kirk")
   .addItem("Iniciar Processos", "showAdminSidebar")
   .addItem("Show ID", "getId")
   .addToUi();
};

function getId() {
  Browser.msgBox('Spreadsheet key: ' + SpreadsheetApp.getActiveSpreadsheet().getId());
};

function showAdminSidebar() {
  const widget = HtmlService.createTemplateFromFile('Index').evaluate();
  SpreadsheetApp.getUi().showSidebar(widget);
};

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
};

function getProcessId() {
  const sheetsId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const responseProcess = UrlFetchApp.fetch(`URL/${sheetsId}`);
  const objectResponseProcess = JSON.parse(responseProcess.getContentText());
  const processId = objectResponseProcess.id;
  const processParentId = objectResponseProcess.parentId;
  return { processId, processParentId };
};

function getSuccessSheetValues() {
  const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const sheetSuccess = allSheets[1];
  const range = sheetSuccess.getDataRange();
  const successSheetValues = range.getValues();
  return { sheetSuccess, successSheetValues };
}

function successHandler() {
  Browser.msgBox('Processo criado com sucesso!');
};

function errorHandler(action) {
  Browser.msgBox(`Ocorreu um erro ao ${action}, tente novamente mais tarde`);
};

function getProcess() {
  SpreadsheetApp.getActiveSheet().getRange('A1:Z1').setValue('');

  try {
    const letterB = 'B';
    const processId = getProcessId().processId;

    const responseFormFields = UrlFetchApp.fetch(`URL${processId}`);
    const kirkFormFields = JSON.parse(responseFormFields.getContentText());
    SpreadsheetApp.getActiveSheet().getRange(`A1`).setValue('Usuário Responsável');
    kirkFormFields.map((field, index) => {
      if (field.label.toLowerCase() === field.name.toLowerCase()) {
        SpreadsheetApp.getActiveSheet().getRange(`${String.fromCharCode(letterB.charCodeAt(0) + index)}1`).setValue(field.label);
      } else {
        SpreadsheetApp.getActiveSheet().getRange(`${String.fromCharCode(letterB.charCodeAt(0) + index)}1`).setValue(`${field.label} | ${field.name}`);
      }
    });
  }
  catch(err) {
    errorHandler('sincronizar');
  }
};

function errorIdk(action){
  const successSheet = getSuccessSheetValues().sheetSuccess;
  const successSheetValuesLength = getSuccessSheetValues().successSheetValues.length;
  successSheet.getRange(`A${successSheetValuesLength + 1}`).setValue(`${action}`).setBackground('#ea4335').setFontColor('white').setFontWeight("bold")
}

function formatDate(date){
  dataAtualFormat = Utilities.formatDate(new Date(),'GMT-03:00',"yyyy-MM-dd")
  var diffEmMilissegundos =  new Date(date).getTime() - new Date(dataAtualFormat).getTime();
  var diffEmDias = Math.floor(diffEmMilissegundos / (1000 * 60 * 60 * 24));
  return diffEmDias
}

function validadeDays (form, processId){
  const {
    'vencimento': vencimento, 
    'vencimento2' : vencimento2,
    'especifique_o_tipo_da_nota':tipo_nf} = form

  if(tipo_nf == '1' || tipo_nf == '2'){
    formatDate(vencimento)
    if(formatDate > 6){
      startProcess(form, processId)
    }else{
      errorIdk('Data < 7d')
    }
  }
  if(tipo_nf == '8'){
    formatDate(vencimento2)
    if(formatDate > 2){   
      startProcess(form, processId)
    } else{
      errorIdk('Data < 3d')
    }
  }
}

function startProcess(form, processId) {
  const folderKey = "CHAVE_PASTA_COMPARTILHADA"; // Chave da pasta compartilhada.
  const {"usuário responsável": responsibleUser, ...data} = form;
  const payload = {
    data,
    username: responsibleUser
  };
  const {
    'link_nf':link_nf, 
    'numero_da_nf_recibo_rpa':numero_NF} = form
  const fileId = getFileIdFromLink(link_nf);

  try{
    if(fileId){
      if(folderKey){
        if(link_nf != 'Arquivo não encontrado'){
          const options = {
            'method': 'post',
            'payload': JSON.stringify(payload),
            'contentType': 'application/json',
          };
          const response = UrlFetchApp.fetch(`URL/${processId}`, options);
          const processInstance = JSON.parse(response.getContentText());
          const idkOfInstance = processInstance.idk;

          const successSheet = getSuccessSheetValues().sheetSuccess;
          const successSheetValuesLength = getSuccessSheetValues().successSheetValues.length;

          const copiedFile = DriveApp.getFileById(fileId)
          const folder = DriveApp.getFolderById(folderKey);
          folder.createFile(copiedFile.getBlob()).setName('IDK'+idkOfInstance+"_"+numero_NF);

          const fileName = "IDK"+idkOfInstance+"_"+numero_NF; // Nome do arquivo a ser verificado
          Utilities.sleep(2000)
          const fileExists = checkFileExistsInFolderByName(folderKey, fileName);

          if (fileExists) {
            successSheet.getRange(`A${successSheetValuesLength + 1}`).setValue(idkOfInstance).setBackground('#34a853').setFontColor('white').setFontWeight("bold");
          } else {
            successSheet.getRange(`A${successSheetValuesLength + 1}`).setValue(idkOfInstance).setBackground('#ea4335').setFontColor('white').setFontWeight("bold");
          } 
        } 
      }
    } else {
      const successSheet = getSuccessSheetValues().sheetSuccess;
      const successSheetValuesLength = getSuccessSheetValues().successSheetValues.length;
      successSheet.getRange(`A${successSheetValuesLength + 1}`).setValue('S/ Anexo').setBackground('#ea4335').setFontColor('white').setFontWeight("bold");
    }
  }catch(error){
    errorIdk('S/ Anexo')
  }
}

function instanceProcess() {
  const values = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  const [columns, ...rows] = values;

  const forms = rows.reduce((result, row) => {
    const formRow = {};
    columns
      .forEach((column, index) => {
        if (column.includes('|')) {
          const fieldName = column.split('|')[1].trim();
          formRow[fieldName.toLowerCase()] = row[index];
        } else {
          formRow[column.toLowerCase()] = row[index];
        }
      });

    result.push(formRow);
    return result;
  }, []);

  const verifyResponsibleField = forms.map(item => item['usuário responsável'] !== '' && true);
  const hasResponsible = !!verifyResponsibleField.length && verifyResponsibleField.every(item => item);

  if (hasResponsible) {
    try {
      const letterB = 'B'
      const processId = getProcessId().processParentId;
      const successSheet = getSuccessSheetValues().sheetSuccess;
      const successSheetValuesLength = getSuccessSheetValues().successSheetValues.length;
      const letterA = 'A'
      forms.forEach((form, indexForm) => {
        validadeDays(form, processId);
        columns.forEach((column, indexColumn) => SpreadsheetApp.getActiveSheet().getRange(`${String.fromCharCode(letterA.charCodeAt(0) + indexColumn)}${indexForm + 2}`).setValue(''));
        }
      );

      //Escrever na aba SUCESSO as informações utilizadas para abrir o chamado
      rows.forEach((row, indexRows) => {
        row.forEach((value, index) => successSheet.getRange(`${String.fromCharCode(letterB.charCodeAt(0) + index)}${indexRows + successSheetValuesLength + 1}`).setValue(`${value}`));
      });
      successHandler();
    }
    catch(err) {
      errorHandler('instanciar');
    }
  } else {
    Browser.msgBox('Por favor, preencha o usuário responsável pela planilha');
  }
};

function checkFileExistsInFolderByName(folderKey, fileName) {
  const folder = DriveApp.getFolderById(folderKey);
  const files = folder.getFilesByName(fileName);
  return files.hasNext(); // Verificar se há algum arquivo com o nome fornecido na pasta
}

function processFolderId(folderId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rangeD = sheet.getRange("J2:J"); // Coluna da NF
    const rangeG = sheet.getRange("T2:T"); // Coluna do Link

    const keywordValues = rangeD.getValues().flat().filter(String);

    const result = keywordValues.map(function(keyword) {
      try {
        var folder = DriveApp.getFolderById(folderId);
      } catch (error) {
        accessDenied("Sem acesso à pasta para buscar as notas fiscais.")
      }

      const files = folder.searchFiles('fullText contains "' + keyword + '" and mimeType = "application/pdf"');

      if (files.hasNext()) {
        const file = files.next();
        const fileId = file.getId();
        const fileLink = "https://drive.google.com/file/d/" + fileId + "/view";
        return [fileLink];
      } else {
        return ["Arquivo não encontrado"];
      }
    });

    rangeG.clearContent();
    rangeG.getSheet().getRange(2, 20, result.length, 1).setValues(result); //Escrevendo aonde é o link 

    return result;
  } catch (error) {
    return [["Erro na execução da função"]];
  }
}

function getFileIdFromLink(link) {
  try {
    var fileId = link.split("/")[5];
    return fileId;
  } catch (e) {
    return null;
  }
}


