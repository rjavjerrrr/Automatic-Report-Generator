function generateEmailHtml() {
  const template = HtmlService.createTemplateFromFile("htmlTemplate"); // Tu plantilla HTML
  const reportData = getReportData();

  template.data = reportData;

  const html = template.evaluate().getContent();
  return html;
}

function sendReportEmail() {
  const recipient = ""; // Cambiar por el correo real
  const subject = "Procurement Pricing PC Report - Weekly Update";
  const htmlBody = generateEmailHtml();

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody,
  });
}

function getReportData() {
  // Obtener datos de la hoja de cálculo
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("datas");
  const dataRange = sheet.getDataRange();
  const rawData = dataRange.getValues();

  // Eliminar encabezados si es necesario
  rawData.shift();

  // Procesar los datos
  const processedData = processRawData(rawData);

  // Crear objeto con la estructura que espera la plantilla
  const reportData = {
    reportDate: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "MMMM d, yyyy"
    ),
    recipientName: "Guillermo Carmona", // Puedes personalizar esto
    completedCount: processedData.completedItems.length,
    inProcessCount: processedData.inProcessItems.length,
    noVolumeCount: processedData.noVolumeItems.length,
    completedItems: processedData.completedItems,
    inProcessItems: processedData.inProcessItems,
    noVolumeItems: processedData.noVolumeItems,
  };

  return reportData;
}

function processRawData(rawData) {
  const completedItems = [];
  const inProcessItems = [];
  const noVolumeItems = [];

  rawData.forEach((row) => {
    const status = row[2]; // Asumiendo que la columna 3 es el estado

    if (status.includes("Completed")) {
      completedItems.push(row);
    } else if (status.includes("In Process")) {
      inProcessItems.push(row);
    } else if (status.includes("No Volume")) {
      noVolumeItems.push(row);
    }
  });

  return {
    completedItems,
    inProcessItems,
    noVolumeItems,
  };
}
