/**
 * Adiciona o menu ao abrir a planilha.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 1º Submenu: Lovable (Itens 1 e 4)
  const menuLovable = ui.createMenu('1. Extração Lovable')
    .addItem('Ordens Oficina Lovable', 'executarProcessoCompleto')
    .addItem('Ordens Oficina Finalizadas', 'extrairValoresFinalizados')
    .addSeparator()
    .addItem('⚡ Atualizar Tudo (Lovable)', 'atualizarTudoLovable');

  // 2º Submenu: Salesforce (Itens 2 e 3)
  const menuSalesforce = ui.createMenu('2. Extração Salesforce')
    .addItem('Serviços Oficina Agenda', 'importSalesforceQueryOficinaAgenda')
    .addItem('Serviços Oficina Custos', 'importarValoresServicosSalesforce')
    .addSeparator()
    .addItem('⚡ Atualizar Tudo (Salesforce)', 'atualizarTudoSalesforce');

  // NOVO Submenu: Extração Geral (Lovable + Salesforce)
  // Posicionado entre o 2 e o 3 conforme solicitado
  const menuGeral = ui.createMenu('3. Extração Geral (Todos)')
    .addItem('🚀 Rodar Todas as Extrações', 'atualizarTudoGeral');

  // 3º Submenu: Enviar Atualizações (Itens 5 e 6)
  const menuAtualizacoes = ui.createMenu('4. Enviar Atualizações')
    .addItem('Atualizar Salesforce', 'updateOrdensSalesforce')
    .addItem('Atualizar Custos (Lovable)', 'updateLovableCustos');

  // Constrói o Menu Principal
  ui.createMenu('🚀 Salesforce Turbi')
    .addSubMenu(menuLovable)
    .addSubMenu(menuSalesforce)
    .addSubMenu(menuGeral) // Menu Geral inserido aqui
    .addSeparator()
    .addSubMenu(menuAtualizacoes)
    .addToUi();
}

// =========================================================================
// FUNÇÕES DE ATUALIZAÇÃO EM MASSA
// =========================================================================

/**
 * Roda TODAS as extrações da planilha (Lovable e Salesforce) em sequência.
 */
function atualizarTudoGeral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Iniciando extração total...', 'Processando', -1);
  
  // Extrações Lovable
  executarProcessoCompleto();
  extrairValoresFinalizados();
  
  // Extrações Salesforce
  importSalesforceQueryOficinaAgenda();
  importarValoresServicosSalesforce();
  
  ss.toast('Todas as extrações foram concluídas com sucesso!', 'Sucesso Geral', 5);
  SpreadsheetApp.getUi().alert('✅ Sucesso! Todas as extrações foram concluídas com sucesso.');
}

// =========================================================================
// FUNÇÕES AUXILIARES PARA OS BOTÕES DE "ATUALIZAR TUDO"
// =========================================================================

/**
 * Roda as extrações do Lovable em sequência.
 */
function atualizarTudoLovable() {
  executarProcessoCompleto();
  extrairValoresFinalizados();
  //SpreadsheetApp.getActiveSpreadsheet().toast('Todas as extrações do Lovable foram concluídas!', 'Sucesso');
  SpreadsheetApp.getUi().alert('✅ Sucesso! Importação Lovable concluída com sucesso.');
}

/**
 * Roda as extrações do Salesforce em sequência.
 */
function atualizarTudoSalesforce() {
  importSalesforceQueryOficinaAgenda();
  importarValoresServicosSalesforce();
  //SpreadsheetApp.getActiveSpreadsheet().toast('Todas as extrações do Salesforce foram concluídas!', 'Sucesso');
  SpreadsheetApp.getUi().alert('✅ Sucesso! Importação SalesForce concluída com sucesso.');
}


function getRequiredScriptProperties(names) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const values = {};
  const missing = [];

  names.forEach((name) => {
    const value = scriptProperties.getProperty(name);
    if (value === null || value === undefined || value.trim() === '') {
      missing.push(name);
    } else {
      values[name] = value;
    }
  });

  if (missing.length > 0) {
    throw new Error(`Propriedades do script ausentes: ${missing.join(', ')}. Configure-as em Apps Script Project Settings > Script Properties antes de executar novamente.`);
  }

  return values;
}

/**
 * Função para conectar com SalesForce
 */
function getSalesforceAccessToken() {
  const properties = getRequiredScriptProperties([
    'SALESFORCE_CLIENT_ID',
    'SALESFORCE_CLIENT_SECRET',
    'SALESFORCE_USERNAME',
    'SALESFORCE_PASSWORD',
    'SALESFORCE_SECURITY_TOKEN'
  ]);
  const payload = {
    'grant_type': 'password',
    'client_id': properties.SALESFORCE_CLIENT_ID,
    'client_secret': properties.SALESFORCE_CLIENT_SECRET,
    'username': properties.SALESFORCE_USERNAME,
    'password': properties.SALESFORCE_PASSWORD + properties.SALESFORCE_SECURITY_TOKEN
  };
  const res = UrlFetchApp.fetch('https://login.salesforce.com/services/oauth2/token', { 'method': 'post', 'payload': payload });
  const json = JSON.parse(res.getContentText());
  return { token: json.access_token, instance: json.instance_url };
}

/**
 * Função para conectar no Lovable
 */
const SUPABASE_URL = "https://tmesktgbvxbjdknontnx.supabase.co";

function executarProcessoCompleto() {
 const token = realizarLoginEAtualizarApoio();
 if (token) {
  importarDadosLovable(token);
 }
}

function getLovableApiKey() {
  return getRequiredScriptProperties(['LOVABLE_API_KEY_ANON']).LOVABLE_API_KEY_ANON;
}

function realizarLoginEAtualizarApoio() {
  const properties = getRequiredScriptProperties([
    'LOVABLE_EMAIL',
    'LOVABLE_PASSWORD',
    'LOVABLE_API_KEY_ANON'
  ]);
  const payload = { email: properties.LOVABLE_EMAIL, password: properties.LOVABLE_PASSWORD };
  const options = {
   method: 'post',
   contentType: 'application/json',
   headers: { apikey: properties.LOVABLE_API_KEY_ANON },
   payload: JSON.stringify(payload),
   muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, options);
  const data = JSON.parse(response.getContentText());
  if (data.access_token) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   let apoioSheet = ss.getSheetByName('Apoio') || ss.insertSheet('Apoio');
   apoioSheet.getRange('A1:B1').setValues([['CHAVE', 'VALOR']]);
   apoioSheet.getRange('A2:B2').setValues([['ACCESS_TOKEN', data.access_token]]);
   return data.access_token;
  }
  return null;
}
