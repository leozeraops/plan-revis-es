/**
 * Buscar Valores reportados no Lovable
 */
function extrairValoresFinalizados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = "Valores_Finalizados";
  let sheet = ss.getSheetByName(nomeAba) || ss.insertSheet(nomeAba);
  
  // =========================================================
  // 1. OBTENDO O TOKEN DINAMICAMENTE
  // =========================================================
  const token = realizarLoginEAtualizarApoio();
  
  if (!token) {
    SpreadsheetApp.getUi().alert('⚠️ Falha de Autenticação: Não foi possível obter o Token do Supabase.');
    return;
  }

  // =========================================================
  // 2. CONFIGURANDO A REQUISIÇÃO PARA A API
  // =========================================================
  const headers = {
    "apikey": getLovableApiKey(),
    "Authorization": "Bearer " + token,
    "Content-Type": "application/json"
  };

  const options = {
    "method": "get",
    "headers": headers,
    "muteHttpExceptions": true
  };

  // Query: Busca ordens finalizadas, trazendo dados do veículo, da oficina e os totais
  // order=scheduled_date.desc -> Traz as mais recentes primeiro
  const url = `${SUPABASE_URL}/rest/v1/maintenance_orders?select=id,sa_number,status,financial_sub_status,total_value,scheduled_date,workshops(name),vehicles(plate,model)&status=eq.finalized&order=scheduled_date.desc`;

  try {
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const ordens = JSON.parse(response.getContentText());
      
      if (!ordens || ordens.length === 0) {
        SpreadsheetApp.getUi().alert("Nenhuma ordem finalizada encontrada.");
        return;
      }

      // =========================================================
      // 3. PREPARANDO OS DADOS PARA A PLANILHA
      // =========================================================
      
      // MODIFICAÇÃO AQUI: Limpa apenas as colunas de A até I (as 9 colunas usadas)
      sheet.getRange("A:I").clear();
      
      // Cabeçalho da Planilha
      const headersPlanilha = [[
        "ID Ordem", "Número SA", "Data Agendada", "Oficina", 
        "Placa", "Modelo", "Status", "Sub-status Financeiro", "Valor Total"
      ]];
      
      // Formatação do Cabeçalho
      sheet.getRange(1, 1, 1, headersPlanilha[0].length)
           .setValues(headersPlanilha)
           .setFontWeight("bold")
           .setBackground("#34a853")
           .setFontColor("white");

      const linhas = [];

      // Loop para extrair as informações do JSON
      ordens.forEach(o => {
        // Tratamento para extrair o nome da oficina
        let nomeOficina = "N/A";
        if (o.workshops) {
          nomeOficina = o.workshops.name || (Array.isArray(o.workshops) && o.workshops[0] ? o.workshops[0].name : "N/A");
        }

        // Formatação rápida de data (de ISO para dd/MM/yyyy)
        let dataAgendada = "";
        if (o.scheduled_date) {
          let dt = new Date(o.scheduled_date);
          dataAgendada = Utilities.formatDate(dt, "GMT-3", "dd/MM/yyyy");
        }

        linhas.push([
          o.id,
          o.sa_number || "S/N",
          dataAgendada,
          nomeOficina,
          o.vehicles?.plate || "N/A",
          o.vehicles?.model || "N/A",
          o.status,
          o.financial_sub_status || "N/A",
          o.total_value || 0
        ]);
      });

      // =========================================================
      // 4. GRAVANDO OS DADOS NA PLANILHA
      // =========================================================
      if (linhas.length > 0) {
        // Escreve tudo de uma vez
        sheet.getRange(2, 1, linhas.length, headersPlanilha[0].length).setValues(linhas);
        
        // Formata a coluna de Valor Total (Coluna I - Índice 9) para Moeda
        sheet.getRange(2, 9, linhas.length, 1).setNumberFormat('R$ #,##0.00');
        
        // Ajusta largura das colunas
        sheet.autoResizeColumns(1, headersPlanilha[0].length);
        
        //SpreadsheetApp.getUi().alert(`✅ Sucesso! ${linhas.length} ordens finalizadas foram importadas.`);
      }

    } else {
      SpreadsheetApp.getUi().alert(`Erro na busca: ${response.getContentText()}`);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Falha na execução: ${e.message}`);
  }
}


/**
 * Outra forma de buscar valores - não utilizado
 */

function extrairValoresLovable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = "Relatorio_Financeiro_Lovable";
  let sheet = ss.getSheetByName(nomeAba) || ss.insertSheet(nomeAba);
  
  // =========================================================
  // 1. OBTENDO O TOKEN DINAMICAMENTE
  // =========================================================
  const token = realizarLoginEAtualizarApoio();
  
  if (!token) {
    SpreadsheetApp.getUi().alert('⚠️ Falha de Autenticação: Não foi possível obter o Token do Supabase.');
    return;
  }

  const headers = {
    "apikey": getLovableApiKey(),
    "Authorization": "Bearer " + token,
    "Content-Type": "application/json"
  };

  const options = {
    "method": "get",
    "headers": headers,
    "muteHttpExceptions": true
  };

  // =========================================================
  // 2. QUERY ATUALIZADA (Trazendo os sub-itens para somar)
  // =========================================================
  // Adicionei 'items' e 'financial_item_details(total_value)' na busca
  const url = `${SUPABASE_URL}/rest/v1/maintenance_orders?select=id,sa_number,status,financial_sub_status,total_value,scheduled_date,workshops(name),vehicles(plate,model),items,financial_item_details(total_value)&or=(status.eq.finalized,financial_sub_status.eq.awaiting_approval)&order=scheduled_date.desc`;

  try {
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const ordens = JSON.parse(response.getContentText());
      
      if (!ordens || ordens.length === 0) {
        SpreadsheetApp.getUi().alert("Nenhuma ordem encontrada com os filtros selecionados.");
        return;
      }

      // =========================================================
      // 3. PREPARANDO OS DADOS PARA A PLANILHA
      // =========================================================
      sheet.clear();
      
      const headersPlanilha = [[
        "ID Ordem", "Número SA", "Data Agendada", "Oficina", 
        "Placa", "Modelo", "Status Geral", "Sub-status Financeiro", "Valor Total"
      ]];
      
      sheet.getRange(1, 1, 1, headersPlanilha[0].length)
           .setValues(headersPlanilha)
           .setFontWeight("bold")
           .setBackground("#4285f4") 
           .setFontColor("white");

      const linhas = [];

      ordens.forEach(o => {
        let nomeOficina = "N/A";
        if (o.workshops) {
          nomeOficina = o.workshops.name || (Array.isArray(o.workshops) && o.workshops[0] ? o.workshops[0].name : "N/A");
        }

        let dataAgendada = "";
        if (o.scheduled_date) {
          let dt = new Date(o.scheduled_date);
          dataAgendada = Utilities.formatDate(dt, "GMT-3", "dd/MM/yyyy");
        }

        // =========================================================
        // CÁLCULO INTELIGENTE DO VALOR (Igual à tela do Lovable)
        // =========================================================
        let valorFinal = parseFloat(o.total_value) || 0;

        // Se a ordem possuir detalhes financeiros na tabela nova, soma tudo
        if (o.financial_item_details && Array.isArray(o.financial_item_details) && o.financial_item_details.length > 0) {
          let somaItensNovos = o.financial_item_details.reduce((acc, curr) => acc + (parseFloat(curr.total_value) || 0), 0);
          if (somaItensNovos > 0) {
            valorFinal = somaItensNovos;
          }
        } 
        // Se for uma ordem mais antiga usando a estrutura de array JSON, soma por ela
        else if (o.items && Array.isArray(o.items) && o.items.length > 0) {
          let somaItensAntigos = o.items.reduce((acc, curr) => acc + (parseFloat(curr.value) || 0), 0);
          if (somaItensAntigos > 0 && valorFinal === 0) {
            valorFinal = somaItensAntigos;
          }
        }

        linhas.push([
          o.id,
          o.sa_number || "S/N",
          dataAgendada,
          nomeOficina,
          o.vehicles?.plate || "N/A",
          o.vehicles?.model || "N/A",
          o.status,
          o.financial_sub_status || "N/A",
          valorFinal // Agora usa o valor calculado e auditado
        ]);
      });

      // =========================================================
      // 4. GRAVANDO OS DADOS E FORMATAÇÃO
      // =========================================================
      if (linhas.length > 0) {
        sheet.getRange(2, 1, linhas.length, headersPlanilha[0].length).setValues(linhas);
        
        // Formata a coluna de Valor Total (Coluna I) para Moeda
        sheet.getRange(2, 9, linhas.length, 1).setNumberFormat('R$ #,##0.00');
        
        // Ajusta largura das colunas
        sheet.autoResizeColumns(1, headersPlanilha[0].length);
        
        //SpreadsheetApp.getUi().alert(`✅ Importação concluída!\nValores zerados foram recalculados com base nas peças.`);
      }

    } else {
      SpreadsheetApp.getUi().alert(`Erro na busca: ${response.getContentText()}`);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Falha na execução: ${e.message}`);
  }
}