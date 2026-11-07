/**
 * Importar dados do Lovable
 */
function importarDadosLovable(token) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 
 // 1. LISTA DOS WORKSHOPS
 /**
  * Para incluir novos centros de serviços o id deve ser atualizado
  */
 const workshopIds = [
   "00b46ecd-4517-43c9-8143-a38f70dfa347",
   "2e78b12e-c3e2-4590-91cd-a4353566109b",
   "e21e0313-e2c3-4fc1-9d6f-0131ff43aae5"
 ];

 const options = {
  "method": "get",
  "headers": {
   "apikey": getLovableApiKey(),
   "Authorization": "Bearer " + token,
   "Content-Type": "application/json"
  },
  "muteHttpExceptions": true
 };

 // 2. PREPARAR AS ABAS ANTES DA BUSCA
 let orderSheet = ss.getSheetByName("OrdensAgendadas") || ss.insertSheet("OrdensAgendadas");
 orderSheet.clear();
 // Cabeçalho com o Workshop na Coluna I
 const orderHeaders = [["Número da SA", "ID Ordem", "Status", "Placa", "Modelo", "Pod/Localização", "Revisão", "Data Agendada", "Workshop"]];
 orderSheet.getRange(1, 1, 1, orderHeaders[0].length).setValues(orderHeaders).setFontWeight("bold").setBackground("#4285f4").setFontColor("white");

 let inspectionSheet = ss.getSheetByName("ItensInspeção") || ss.insertSheet("ItensInspeção");
 inspectionSheet.clear();
 const inspHeaders = [["ID da Ordem", "Nome do Item", "Classe (Tipo)"]];
 inspectionSheet.getRange(1, 1, 1, inspHeaders[0].length).setValues(inspHeaders).setFontWeight("bold").setBackground("#34a853").setFontColor("white");

 const orderRows = [];
 const inspectionRows = [];
 let sucessoNasBuscas = false;

 // 3. BUSCA E JUNTA OS DADOS NO MESMO FORMATO QUE O SEU FUNCIONAL
 workshopIds.forEach(id => {
  const url = `${SUPABASE_URL}/rest/v1/maintenance_orders?select=*,workshops(name),vehicles(*),booking_tokens(client_name)&workshop_id=eq.${id}&status=neq.finalized&order=scheduled_date.asc`;
  
   try {
     const response = UrlFetchApp.fetch(url, options);
     if (response.getResponseCode() === 200) {
       const orders = JSON.parse(response.getContentText());
       
       if (orders && orders.length > 0) {
         sucessoNasBuscas = true;
         
         orders.forEach(o => {
           // Proteção: Garante que extraímos apenas o texto do nome do workshop
           let nomeDoWorkshop = "N/A";
           if (o.workshops) {
             nomeDoWorkshop = o.workshops.name || (Array.isArray(o.workshops) && o.workshops[0] ? o.workshops[0].name : "N/A");
           }

           orderRows.push([
             o.sa_number || "S/N",
             o.id,
             o.status,
             o.vehicles?.plate || o.plate,
             o.vehicles?.model || o.model,
             o.vehicles?.pod_name || "N/A", 
             o.revision_label || "N/A",
             o.scheduled_date,
             nomeDoWorkshop // Inserido na Coluna I
           ]);

           if (o.items && Array.isArray(o.items)) {
             o.items.forEach(item => {
               inspectionRows.push([
                 o.id,
                 item.item_name,
                 item.item_class
               ]);
             });
           }
         });
       }
     }
   } catch (e) {
     Logger.log("Falha ao processar o ID: " + id);
   }
 });

 if (!sucessoNasBuscas) {
  SpreadsheetApp.getUi().alert("Nenhuma ordem encontrada nos workshops.");
  return;
 }

 // 4. GRAVA TODOS OS DADOS NAS PLANILHAS DE UMA SÓ VEZ
 if (orderRows.length > 0) {
   orderSheet.getRange(2, 1, orderRows.length, orderHeaders[0].length).setValues(orderRows);
   orderSheet.autoResizeColumns(1, orderHeaders[0].length);
 }

 if (inspectionRows.length > 0) {
   inspectionSheet.getRange(2, 1, inspectionRows.length, inspHeaders[0].length).setValues(inspectionRows);
   inspectionSheet.autoResizeColumns(1, inspHeaders[0].length);
 }

 //SpreadsheetApp.getUi().alert("Importação concluída: " + orderRows.length + " ordens e " + inspectionRows.length + " itens de inspeção.");
}

