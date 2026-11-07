/**
 * Busca os serviços (Servico__c) do Salesforce e limpa APENAS as colunas de dados.
 */
function importarValoresServicosSalesforce() {
  const auth = getSalesforceAccessToken();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DadosServicos') || ss.insertSheet('DadosServicos');
  
  const soql = "SELECT Ordem_de_Manutencao__r.Id, CreatedDate,  Ordem_de_Manutencao__r.WorkOrderNumber, Tipo__c, Tipo_de_Insumo__c, valor_estimado__c, Quantidade__c, Valor_Total_estimado__c,        Ordem_de_Manutencao__r.CA_DetalheDoServico__c FROM Servico__c WHERE Ordem_de_Manutencao__r.CA_DetalheDoServico__c = 'Revisão Interna' AND Ordem_de_Manutencao__r.CreatedDate >= LAST_N_DAYS:30 AND Ordem_de_Manutencao__r.CreatedDate <= NEXT_N_DAYS:30 AND Ordem_de_Manutencao__r.Status != 'Cancelado'";
  
  let allRecords = [];
  let nextRecordsUrl = "/services/data/v57.0/query?q=" + encodeURIComponent(soql);

  try {
    do {
      const response = UrlFetchApp.fetch(auth.instance + nextRecordsUrl, {
        "method": "get",
        "headers": { "Authorization": "Bearer " + auth.token }
      });
      
      const result = JSON.parse(response.getContentText());
      allRecords = allRecords.concat(result.records);
      nextRecordsUrl = result.nextRecordsUrl; 
      
    } while (nextRecordsUrl);

    if (allRecords.length > 0) {
      const rows = allRecords.map(function(record) {
        return [
          record.Ordem_de_Manutencao__r ? record.Ordem_de_Manutencao__r.Id : "",
          record.CreatedDate,
          record.Ordem_de_Manutencao__r ? record.Ordem_de_Manutencao__r.WorkOrderNumber : "",
          record.Tipo__c || "",
          record.Tipo_de_Insumo__c || "",
          record.valor_estimado__c || 0,
          record.Quantidade__c || 0,
          record.Valor_Total_estimado__c || 0,
          record.Ordem_de_Manutencao__r ? record.Ordem_de_Manutencao__r.CA_DetalheDoServico__c : ""
        ];
      });

      // --- SEÇÃO DE LIMPEZA INTELIGENTE ---
      // Definimos o cabeçalho (9 colunas)
      const headers = [["ID OM", "Data Criação", "Número OM", "Tipo", "Insumo", "Valor Unit.", "Qtd", "Total Estimado", "Detalhe Serviço"]];
      
      // Limpa apenas as colunas A até I (1 a 9) preservando o resto da planilha
      if (sheet.getLastRow() > 0) {
        // Limpa o conteúdo das 9 colunas onde os dados entram, mantendo formatação e fórmulas fora desse range
        sheet.getRange(2, 1, Math.max(sheet.getLastRow(), 2), headers[0].length).clearContent();
      }
      
      // Grava os cabeçalhos e os novos dados
      sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight("bold").setBackground("#cfe2f3");
      sheet.getRange(2, 1, rows.length, headers[0].length).setValues(rows);
      
      // Ajusta largura apenas das colunas importadas
      for (let i = 1; i <= headers[0].length; i++) {
        sheet.autoResizeColumn(i);
      }

      //SpreadsheetApp.getUi().alert("✅ Sucesso: " + allRecords.length + " serviços importados sem apagar suas fórmulas!");
    } else {
      SpreadsheetApp.getUi().alert("⚠️ Nenhum serviço encontrado nos últimos 90 dias.");
    }

  } catch (e) {
    Logger.log("Erro: " + e.message);
    SpreadsheetApp.getUi().alert("❌ Falha na busca: " + e.message);
  }
}