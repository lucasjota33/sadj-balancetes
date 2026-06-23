// ============================================================
//  SADJ BALANCETES — Google Apps Script Backend
// ============================================================

const CONFIG = {
  PASTA_RAIZ_ID: "1AB0fgS-BI9Qdb6J0EN3TzjYGmTBhs7Ir",
  PLANILHA_ID: "1QDdrJv7vRWckLojm8v6UFX7ORFDIRjcrUffUvZevmoU", // Sua planilha oficial
  MESES: ["", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]
};

function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    let resultado;
    
    if (dados.acao === "uploadComprovante") {
      resultado = uploadComprovante(dados);
    } else if (dados.acao === "salvarBalancete") {
      resultado = salvarBalancete(dados);
    } else if (dados.acao === "listarPastas") {
      resultado = listarPastas();
    } else {
      resultado = { erro: "Ação desconhecida" };
    }
    return ContentService.createTextOutput(JSON.stringify(resultado)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ erro: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) { return ContentService.createTextOutput(JSON.stringify({ status: "SADJ Ativo" })).setMimeType(ContentService.MimeType.JSON); }

function obterPastaMes(mes, ano) {
  const nomePasta = mes.toUpperCase() + "_" + String(ano).slice(-2);
  const pastaRaiz = DriveApp.getFolderById(CONFIG.PASTA_RAIZ_ID);
  const pastas = pastaRaiz.getFoldersByName(nomePasta);
  return pastas.hasNext() ? pastas.next() : pastaRaiz.createFolder(nomePasta);
}

function obterSubpasta(pastaPai, nomeSubpasta) {
  const subs = pastaPai.getFoldersByName(nomeSubpasta);
  return subs.hasNext() ? subs.next() : pastaPai.createFolder(nomeSubpasta);
}

function uploadComprovante(dados) {
  // Aceita pasta enviada pelo frontend OU a configurada no script
  const pastaRaizId = (dados.pastaId && dados.pastaId.trim()) ? dados.pastaId.trim() : CONFIG.PASTA_RAIZ_ID;
  const pastaRaiz = DriveApp.getFolderById(pastaRaizId);

  // Nome do mês sem acento, ex: MARCO_26
  const mesNorm = dados.mes.toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "_");
  const nomeMes = mesNorm + "_" + String(dados.ano).slice(-2);
  const pastaMes = obterSubpasta(pastaRaiz, nomeMes);
  const pastaComp = obterSubpasta(pastaMes, "COMPROVANTES");
  const pastaTipo = obterSubpasta(pastaComp, dados.tipo === "despesa" ? "DESPESAS" : "RECEITAS");

  const blob = Utilities.newBlob(Utilities.base64Decode(dados.base64), dados.mimeType, dados.nomeArquivo);
  const arquivo = pastaTipo.createFile(blob);

  // Qualquer pessoa com o link pode visualizar
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = arquivo.getId();

  return {
    sucesso: true,
    fileId: fileId,
    url: arquivo.getUrl(),
    // /view abre no Google Drive Viewer no navegador (funciona para PDF, imagem, etc.)
    urlPreview: "https://drive.google.com/file/d/" + fileId + "/view",
    // Para imagens: renderiza diretamente sem precisar do Viewer
    urlDirect: "https://drive.google.com/uc?export=view&id=" + fileId,
    nome: arquivo.getName()
  };
}

function salvarBalancete(dados) {
  const planilha = SpreadsheetApp.openById(CONFIG.PLANILHA_ID);
  const nomeMes = dados.mes.toUpperCase() + "_" + String(dados.ano).slice(-2);
  
  let aba = planilha.getSheetByName(nomeMes);
  if (!aba) {
    aba = planilha.insertSheet(nomeMes);
  }
  
  aba.clearContents();
  escreverBalancete(aba, dados);
  const urlPdf = salvarPdfNoDrive(dados);
  
  return { sucesso: true, urlPlanilha: planilha.getUrl(), urlPdf: urlPdf };
}

function escreverBalancete(aba, d) {
  aba.getRange("B2").setValue("Diretoria Financeira SADJ 2026 - PRESTAÇÃO DE CONTAS");
  aba.getRange("K5").setValue("SADJ " + d.mes + "/" + d.ano);
  aba.getRange("B5").setValue("SALDO INICIAL");
  aba.getRange("G5").setValue(parseFloat(d.saldoInicial));

  // Lógica simplificada para inserção (pode ser refinada conforme o seu template exato)
  aba.getRange("B17").setValue("DESPESAS");
  d.despesas.forEach((dep, i) => {
    const lin = 20 + i;
    aba.getRange("B" + lin).setValue(i + 1);
    aba.getRange("C" + lin).setValue(dep.data);
    aba.getRange("E" + lin).setValue(parseFloat(dep.valor || 0));
    aba.getRange("G" + lin).setValue(dep.descricao);
    if (dep.urlDrive) {
      aba.getRange("M" + lin).setFormula('=HYPERLINK("' + dep.urlDrive + '","' + dep.comprovante + '")');
    } else {
      aba.getRange("M" + lin).setValue(dep.comprovante);
    }
  });

  // Saldo Final e Assinaturas
  const linAsn = 50; 
  aba.getRange("B" + linAsn).setValue("PRESIDENTE DA SADJ\n" + d.presidente);
  aba.getRange("D" + linAsn).setValue("TESOUREIRA DA SADJ\n" + d.tesoureira);
  SpreadsheetApp.flush();
}

function salvarPdfNoDrive(dados) {
  if (!dados.pdfBase64) return null;
  const pastaMes = obterPastaMes(dados.mes, dados.ano);
  const blob = Utilities.newBlob(Utilities.base64Decode(dados.pdfBase64), "application/pdf", "Balancete_SADJ_" + dados.mes + "_" + dados.ano + ".pdf");
  const arquivo = pastaMes.createFile(blob);
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return arquivo.getUrl();
}

function listarPastas() {
  const pastas = DriveApp.getFolderById(CONFIG.PASTA_RAIZ_ID).getFolders();
  const res = [];
  while (pastas.hasNext()) res.push({ nome: pastas.next().getName() });
  return { pastas: res };
}