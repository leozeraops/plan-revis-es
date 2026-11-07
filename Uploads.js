/**
 * Atualizar os custos do Lovable
 */
function updateLovableCustos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pendentes Custos');
  const ui = SpreadsheetApp.getUi();

  // ==========================================
  // 1. SOBE O POPUP DE "EM ANDAMENTO"
  // ==========================================
  const htmlCarregando = HtmlService.createHtmlOutput(
    '<div style="text-align: center; font-family: Arial, sans-serif; padding: 20px;">' +
    '<h3>Processando...</h3>' +
    '<p>Enviando atualizações para o Lovable.<br>Isso pode levar alguns segundos.</p>' +
    '</div>'
  ).setWidth(350).setHeight(150);
  
  ui.showModelessDialog(htmlCarregando, '⏳ Atualização em andamento');

  if (!sheet) {
    ui.alert('⚠️ Aba "Pendentes Custos" não encontrada!');
    return;
  }

  // =========================================================
  // LIMPEZA INICIAL DA COLUNA U (Status)
  // A coluna Status agora é a U (Índice 21 no getRange)
  // =========================================================
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 21, lastRow - 1, 1).clearContent();
  }

  // =========================================================
  // OBTENDO O TOKEN DINAMICAMENTE
  // =========================================================
  const token = realizarLoginEAtualizarApoio();
  if (!token) {
    ui.alert('⚠️ Falha de Autenticação: Não foi possível obter o Token.');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  let ordensAtualizadas = 0;
  let ordensComErro = 0;
  
  // Matriz de memória para a Coluna U
  let statusOutput = [];
  statusOutput.push(["Status API"]); // Cabeçalho da Coluna U

  const headersAPI = { 
    'apikey': getLovableApiKey(),             
    'Authorization': 'Bearer ' + token, 
    'Content-Type': 'application/json',
    'Prefer': 'return=minimal' 
  };
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Mapeamento Atualizado (A=0, B=1, C=2...)
    const idOrdem    = row[2] ? String(row[2]).trim() : "";  // Coluna C (ID)
    const permissaoT = row[19];                              // Coluna T (Valid - TRUE/FALSE)
    
    // =========================================================
    // FILTRO: SÓ RODA SE TIVER ID E COLUNA T FOR TRUE
    // =========================================================
    if (idOrdem && (permissaoT === true || String(permissaoT).toUpperCase() === "TRUE")) {
      
      const valorTotal         = parseFloat(row[18]) || 0; // Coluna S (Valor Total)
      
      const vOleo              = parseFloat(row[8])  || 0; // Coluna I
      const vFiltroOleo        = parseFloat(row[9])  || 0; // Coluna J
      const vFiltroComb        = parseFloat(row[10]) || 0; // Coluna K
      const vPastilhaDianteira = parseFloat(row[11]) || 0; // Coluna L
      const vFiltroArMotor     = parseFloat(row[12]) || 0; // Coluna M
      const vPastilhaTraseira  = parseFloat(row[13]) || 0; // Coluna N
      const vMaoDeObra         = parseFloat(row[14]) || 0; // Coluna O
      const vDiscos            = parseFloat(row[15]) || 0; // Coluna P (NOVO)
      const vAditivoRadiador   = parseFloat(row[16]) || 0; // Coluna Q 
      
      try {
        // CHAMADA 1: PATCH
        const payloadPatch = {
          "total_value": valorTotal,
          "items": [
            {"item_name": "Filtro de Combustivel", "item_class": "mandatory", "value": vFiltroComb},
            {"item_name": "Filtro de Óleo", "item_class": "mandatory", "value": vFiltroOleo},
            {"item_name": "Óleo", "item_class": "mandatory", "value": vOleo},
            {"item_name": "Pastilha dianteira", "item_class": "mandatory", "value": vPastilhaDianteira},
            {"item_name": "Aditivo de Radiador", "item_class": "inspection", "value": vAditivoRadiador}, 
            {"item_name": "Bomba de Combustível", "item_class": "inspection", "value": 0},
            {"item_name": "Correia", "item_class": "inspection", "value": 0},
            {"item_name": "Discos", "item_class": "inspection", "value": vDiscos}, // ATUALIZADO
            {"item_name": "Filtro de Ar Motor", "item_class": "inspection", "value": vFiltroArMotor},
            {"item_name": "Filtro do ar condicionado", "item_class": "inspection", "value": 0},
            {"item_name": "Fluido da embreagem e freio", "item_class": "inspection", "value": 0},
            {"item_name": "Pastilha traseira", "item_class": "inspection", "value": vPastilhaTraseira},
            {"item_name": "Velas", "item_class": "inspection", "value": 0}
          ],
          "extra_items": [],
          "supplier_values_locked": true,
          "status": "finalized",
          "financial_sub_status": "awaiting_approval"
        };
        
        let resPatch = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/maintenance_orders?id=eq.${idOrdem}`, {
          'method': 'patch', 'headers': headersAPI, 'payload': JSON.stringify(payloadPatch), 'muteHttpExceptions': true
        });
        if (resPatch.getResponseCode() !== 204 && resPatch.getResponseCode() !== 200) throw new Error(`Erro PATCH: ${resPatch.getContentText()}`);

        // CHAMADA 2: DELETE
        let resDelete = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/financial_item_details?maintenance_order_id=eq.${idOrdem}`, {
          'method': 'delete', 'headers': headersAPI, 'muteHttpExceptions': true
        });
        if (resDelete.getResponseCode() !== 204 && resDelete.getResponseCode() !== 200) throw new Error(`Erro DELETE`);

        // CHAMADA 3: POST
        let payloadPost = [];
        const adicionarItemPost = (nome, categoria, origem, valor) => {
          if (valor > 0) payloadPost.push({ "maintenance_order_id": idOrdem, "item_name": nome, "item_category": categoria, "item_source": origem, "source_id": null, "unit_value": valor, "quantity": 1, "total_value": valor });
        };

        adicionarItemPost("Óleo", "obrigatorio", "matrix", vOleo);
        adicionarItemPost("Filtro de Óleo", "obrigatorio", "matrix", vFiltroOleo);
        adicionarItemPost("Filtro de Combustivel", "obrigatorio", "matrix", vFiltroComb);
        adicionarItemPost("Pastilha dianteira", "obrigatorio", "matrix", vPastilhaDianteira);
        adicionarItemPost("Filtro de Ar Motor", "obrigatorio", "matrix", vFiltroArMotor);
        adicionarItemPost("Pastilha traseira", "obrigatorio", "matrix", vPastilhaTraseira);
        adicionarItemPost("Discos", "inspecao", "matrix", vDiscos); // NOVO ITEM
        adicionarItemPost("Aditivo de Radiador", "inspecao", "matrix", vAditivoRadiador); 
        
        adicionarItemPost("Mão de Obra", "mao_de_obra", "manual", vMaoDeObra);

        if (payloadPost.length > 0) {
          let resPost = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/financial_item_details?maintenance_order_id=eq.${idOrdem}`, {
            'method': 'post', 'headers': headersAPI, 'payload': JSON.stringify(payloadPost), 'muteHttpExceptions': true
          });
          if (resPost.getResponseCode() !== 201 && resPost.getResponseCode() !== 200 && resPost.getResponseCode() !== 204) throw new Error(`Erro POST`);
        }

        statusOutput.push(["200 Sucesso"]);
        ordensAtualizadas++;

      } catch (e) {
        statusOutput.push([`Falha: ${e.message}`]);
        ordensComErro++;
      }
    } else {
      // Se a coluna T estiver FALSE ou o ID vazio, apenas mantém a célula vazia na memória
      statusOutput.push([""]); 
    }
  }
  
  // Escreve tudo de uma vez na Coluna U (Índice 21 para o getRange)
  sheet.getRange(1, 21, statusOutput.length, 1).setValues(statusOutput);

  // ==========================================
  // FECHA O POPUP E DÁ A MENSAGEM FINAL
  // ==========================================

  // Roda a próxima função automaticamente após clicar no OK
  executarProcessoCompleto();
  extrairValoresFinalizados();
  
  // Esse comando cria uma mini-tela invisível que manda o Google fechar o popup de carregamento
  const htmlFechar = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
  ui.showModelessDialog(htmlFechar, 'Concluído'); 
  
  // O seu alerta final tradicional (O script pausa aqui esperando o "OK")
  ui.alert(`✅ Processo Concluído!\n\nAtualizadas: ${ordensAtualizadas}\nErros: ${ordensComErro}`);
}

/**
 * Atualizar Ordens do SalesForce
 */
function updateOrdensSalesforce() {
  const auth = getSalesforceAccessToken();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('LovableAgendadas');
  const sheetApoio = ss.getSheetByName('Apoio');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠️ Aba "LovableAgendadas" não encontrada!');
    return;
  }
  if (!sheetApoio) {
    SpreadsheetApp.getUi().alert('⚠️ Aba "Apoio" não encontrada! Necessária para buscar o ID do Local.');
    return;
  }
  
  // ==========================================
  // CARREGAR DADOS DA ABA APOIO (Dicionário para Lookup Rápido)
  // ==========================================
  const dataApoio = sheetApoio.getDataRange().getValues();
  let mapaLocalServico = {};
  
  for (let i = 1; i < dataApoio.length; i++) {
    // Aba Apoio: Coluna M = Índice 12 (Chave de busca), Coluna N = Índice 13 (ID do Local)
    const chaveBusca = dataApoio[i][12]; 
    const idLocal = dataApoio[i][13];
    
    if (chaveBusca) {
      mapaLocalServico[chaveBusca] = idLocal;
    }
  }

  // ==========================================
  // PROCESSAR ABA LOVABLE AGENDADAS
  // ==========================================
  const data = sheet.getDataRange().getValues();
  let ordensAtualizadas = 0;
  let ordensComErro = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Mapeamento
    const chaveParaBusca = row[8]; // Coluna I: O que vamos usar para buscar na aba Apoio (Placa, etc.)
    const numeroRevisao  = row[9]; // Coluna J: Número da revisão
    const idOM           = row[11];// Coluna L: ID do Salesforce (Ex: 0WONp000009Ke7NOAS)
    const om             = row[12];// Coluna M: Número da OM (Ex: 00012345)
    const statusP        = row[15];// Coluna P: Status de processamento (TRUE/FALSE - via ArrayFormula)
    
    // Condição de execução: Se P for FALSE e a Ordem não for vazia
    if ((statusP === false || statusP === "FALSE" || statusP === "") && om !== "") {
      
      // Busca o ID do Local no dicionário criado a partir da aba Apoio
      const localServicoID = mapaLocalServico[chaveParaBusca] || "";
      
      // Se não encontrou o ID na aba Apoio, já barra e avisa na Coluna Q
      if (localServicoID === "") {
        sheet.getRange(i + 1, 17).setValue(`Erro: Local de serviço não encontrado na aba Apoio para '${chaveParaBusca}'`);
        ordensComErro++;
        continue; // Pula para a próxima linha sem chamar a API
      }

      const payload = {
        "Numero_da_Revisao__c": numeroRevisao,
        "ServiceLocation__c": localServicoID // Agora enviamos o ID real
      };
      
      const url = `${auth.instance}/services/data/v57.0/sobjects/WorkOrder/${idOM}`;
      const options = {
        'method': 'patch',
        'contentType': 'application/json',
        'headers': { 'Authorization': 'Bearer ' + auth.token },
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };
      
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode === 204 || responseCode === 200) {
          // Apenas gravamos o sucesso na Coluna Q. O Array da Coluna P faz o resto!
          sheet.getRange(i + 1, 17).setValue("200 Sucesso");
          ordensAtualizadas++;
        } else {
          const erroMsg = `Erro ${responseCode}: ${response.getContentText()}`;
          sheet.getRange(i + 1, 17).setValue(erroMsg);
          Logger.log(`Erro na OM ${om}: ${erroMsg}`);
          ordensComErro++;
        }
      } catch (e) {
        sheet.getRange(i + 1, 17).setValue(`Falha de Execução: ${e.message}`);
        Logger.log(`Falha na requisição da OM ${om}: ${e.message}`);
        ordensComErro++;
      }
    }
  }  
  
  // Exibe o alerta na tela e "pausa" o script até você clicar em OK
  SpreadsheetApp.getUi().alert(`✅ Processo Finalizado!\n\nAtualizadas: ${ordensAtualizadas}\nErros: ${ordensComErro}\n\nConfira a Coluna Q para detalhes.`);
  
  // Assim que clicar em OK, ele roda a função abaixo automaticamente:
  importSalesforceQueryOficinaAgenda(); 
}