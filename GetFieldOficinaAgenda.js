/**
 * Buscar Ordens do SalesForce
 */
function importSalesforceQueryOficinaAgenda() {
  const auth = getSalesforceAccessToken();
  
  // =======================================================
  // PASSO 1: Buscar WorkOrders e seus ServiceAppointments
  // =======================================================
  const soqlWO = `
    SELECT Id, WorkOrderNumber, CreatedDate, CA_DetalheDoServico__c, CA_TipoDeServico__c, 
           Numero_da_Revisao__c, Placa__c, Status, Priority, WorkType.Name, 
           Account.Name, Contact.Name, ServiceLocation__r.Name,
           (SELECT Id, AppointmentNumber, Status, SchedStartTime, SchedEndTime, 
                   ActualStartTime, ActualEndTime, DueDate, ServiceTerritory.Name 
            FROM ServiceAppointments)
    FROM WorkOrder
    WHERE CA_DetalheDoServico__c = 'Revisão Interna' 
    AND CreatedDate >= LAST_N_DAYS:30 
    AND CreatedDate <= NEXT_N_DAYS:30
    AND Status != 'Cancelado'
  `;

  let urlWO = `${auth.instance}/services/data/v57.0/query/?q=${encodeURIComponent(soqlWO)}`;
  let resWO = UrlFetchApp.fetch(urlWO, {
    'method': 'get',
    'headers': { 'Authorization': 'Bearer ' + auth.token }
  });
  
  const resultWO = JSON.parse(resWO.getContentText());

  let saIds = [];
  let todasOrdens = [];

  if (resultWO.records && resultWO.records.length > 0) {
    resultWO.records.forEach(wo => {
      todasOrdens.push(wo);
      if (wo.ServiceAppointments && wo.ServiceAppointments.records && wo.ServiceAppointments.records.length > 0) {
        wo.ServiceAppointments.records.forEach(sa => {
          if (sa.Id) saIds.push(sa.Id);
        });
      }
    });
  }

  if (todasOrdens.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ Nenhuma ordem encontrada para esses filtros.');
    return;
  }

  // =======================================================
  // PASSO 2: Buscar os Nomes dos Recursos (Técnicos)
  // =======================================================
  let mapaRecursos = {}; 
  if (saIds.length > 0) {
    const TAMANHO_LOTE = 40; 
    for (let i = 0; i < saIds.length; i += TAMANHO_LOTE) {
      let loteIds = saIds.slice(i, i + TAMANHO_LOTE);
      let soqlAR = `SELECT ServiceAppointmentId, ServiceResource.Name FROM AssignedResource WHERE ServiceAppointmentId IN ('${loteIds.join("','")}')`;
      let urlAR = `${auth.instance}/services/data/v57.0/query/?q=${encodeURIComponent(soqlAR)}`;
      let resAR = UrlFetchApp.fetch(urlAR, {
        'method': 'get',
        'headers': { 'Authorization': 'Bearer ' + auth.token }
      });
      let resultAR = JSON.parse(resAR.getContentText());
      if (resultAR.records) {
        resultAR.records.forEach(ar => {
          mapaRecursos[ar.ServiceAppointmentId] = ar.ServiceResource ? ar.ServiceResource.Name : "";
        });
      }
    }
  }

  // =======================================================
  // PASSO 3: Montar os dados Finais e Gravar na Planilha
  // =======================================================
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OficinaAgenda') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('OficinaAgenda');
  
  const headers = [[
    "Recurso", "Id_SA", "SA", "Data Criação OM", "Tipo de ordem", "Detalhe da ordem", 
    "Número da revisão", "Local de serviço", "Placa", "Nome", "Id_OM", 
    "OM", "StatusOM", "Inicio Real", "Fim Real", "Inicio Agendamento", 
    "Vencimento", "Território", "StatusSA"
  ]];

  let finalData = [];

  const formatarData = (dataIso) => {
    if (!dataIso) return "";
    return Utilities.formatDate(new Date(dataIso), "GMT-3", "dd/MM/yyyy HH:mm:ss");
  };

  todasOrdens.forEach(wo => {
    // Caso a Ordem TENHA agendamentos
    if (wo.ServiceAppointments && wo.ServiceAppointments.records && wo.ServiceAppointments.records.length > 0) {
      wo.ServiceAppointments.records.forEach(sa => {
        let nomeRecurso = mapaRecursos[sa.Id] || "";

        finalData.push([
          nomeRecurso,
          sa.Id || "",
          sa.AppointmentNumber || "",
          formatarData(wo.CreatedDate),
          wo.WorkType ? wo.WorkType.Name : "",
          wo.CA_DetalheDoServico__c || "",
          wo.Numero_da_Revisao__c || "",
          wo.ServiceLocation__r ? wo.ServiceLocation__r.Name : "",
          wo.Placa__c || "",
          wo.Contact ? wo.Contact.Name : (wo.Account ? wo.Account.Name : ""),
          wo.Id || "",
          wo.WorkOrderNumber || "",
          wo.Status || "",
          formatarData(sa.ActualStartTime),
          formatarData(sa.ActualEndTime),
          formatarData(sa.SchedStartTime),
          formatarData(sa.DueDate),
          sa.ServiceTerritory ? sa.ServiceTerritory.Name : "",
          sa.Status || ""
        ]);
      });
    } 
    // Caso a Ordem NÃO TENHA agendamento
    else {
      finalData.push([
        "", "", "", 
        formatarData(wo.CreatedDate),
        wo.WorkType ? wo.WorkType.Name : "",
        wo.CA_DetalheDoServico__c || "",
        wo.Numero_da_Revisao__c || "",
        wo.ServiceLocation__r ? wo.ServiceLocation__r.Name : "",
        wo.Placa__c || "",
        wo.Contact ? wo.Contact.Name : (wo.Account ? wo.Account.Name : ""),
        wo.Id || "",
        wo.WorkOrderNumber || "",
        wo.Status || "",
        "", "", "", "", "", ""
      ]);
    }
  });

  // =======================================================
  // MODIFICAÇÃO AQUI: Limpa apenas as colunas de A até S
  // =======================================================
  sheet.getRange("A:S").clearContent(); 
  
  sheet.getRange(1, 1, 1, headers[0].length)
       .setValues(headers)
       .setFontWeight("bold")
       .setBackground("#cfe2f3");
  
  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, finalData[0].length).setValues(finalData);
    sheet.setFrozenRows(1);
    for (let i = 1; i <= headers[0].length; i++) { 
      sheet.autoResizeColumn(i); 
    }
    //SpreadsheetApp.getUi().alert('✅ Sucesso! Importação concluída com as datas de criação das OMs.');
  }
}