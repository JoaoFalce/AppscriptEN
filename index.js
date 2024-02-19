let app = SpreadsheetApp;
let spreadsheet = app.getActiveSpreadsheet();
let sheet = spreadsheet.getSheetByName("DeclaracaoVisto");
//Declara variáveis para as planilhas e pega a planilha ativa pelo nome


let termoId = ""
//ID da planilha no Google Sheets



// Função que gera os termos de prorrogação em Docs
function gerarDeclaracaoVisto() {
  let matriculas = sheet.getRange("A4:A14").getDisplayValues();
  let nomes = sheet.getRange("B4:B14").getDisplayValues();
  let cargos = sheet.getRange("C4:C14").getDisplayValues();
  let rgs = sheet.getRange("D4:D14").getDisplayValues();
  let cpfs = sheet.getRange("E4:E14").getDisplayValues();
  let cnpj = sheet.getRange("G4:G14").getDisplayValues();
  let empresas = sheet.getRange("F4:F14").getDisplayValues();
  let dias = sheet.getRange("I4:I14").getDisplayValues();
  let meses = sheet.getRange("J4:J14").getDisplayValues();
  let anos = sheet.getRange("K4:K14").getDisplayValues();
  let mesemingles = sheet.getRange("L4:L14").getDisplayValues();
  let tipos = sheet.getRange("M4:M14").getDisplayValues();
  let diashoje = sheet.getRange("N4:N14").getDisplayValues();
  let meseshoje = sheet.getRange("O4:O14").getDisplayValues();
  let anoshoje = sheet.getRange("P4:P14").getDisplayValues();
  //LET seta as variáveis de acordo com a posição delas na planilha, por exemplo, MATRICULAS está em A4 até A14.
  
  for (let index = 0; index < nomes.length; index++) {
    if (matriculas[index][0] !== "") {
      // Se matrículas for igual a NADA, ou seja, se não tiver nada escrito na planilha, NÃO irá imprimir o termo de prorrogação.
      
      
      let parametros = {
        nome: nomes[index][0], 
        cargo: cargos[index][0],
        rg: rgs[index][0],
        cpf: cpfs[index][0],
        matricula: matriculas[index][0],
        cnpj: cnpj[index][0],
        empresa: empresas[index][0],
        dia: dias[index][0],
        mes: meses[index][0],
        ano: anos[index][0],
        mesingles: mesemingles[index][0],
        tipo: tipos[index][0],
        diahoje: diashoje[index][0],
        meshoje: meseshoje[index][0],
        anohoje: anoshoje[index][0],
        
      // Foram criadas propriedades dentro do objeto PARÂMETROS e foi atribuibo o valor do elemento no índice "index" dos arrays (nome, cargo, etc) na posição 0, ou seja, a primeira posição na lista.
      //
      }
       
  let termoDocId = DriveApp.getFileById(termoId).makeCopy(parametros.matricula + " - " + parametros.nome + " - Declaração Visto EN-US").getId();
  // seta o nome do arquivo e cria uma cópia.
  let termoDoc = DocumentApp.openById(termoDocId);
               
      termoDoc.replaceText(
      '{NOME}', parametros.nome
      ),
      termoDoc.replaceText(
      '{CARGO}', parametros.cargo
      ),
      termoDoc.replaceText(
      '{RG}', parametros.rg
      ),
      termoDoc.replaceText(
      '{CPF}', parametros.cpf
      ),
      termoDoc.replaceText(
      '{MATRICULA}', parametros.matricula
      ),
      termoDoc.replaceText(
      '{00.000.000/0000-00}', parametros.cnpj
      ),
      termoDoc.replaceText(
      '{EMPRESA}', parametros.empresa
      ),
      termoDoc.replaceText(
      '{DD}', parametros.dia
      ),
      termoDoc.replaceText(
      '{MM}', parametros.mes
      ),
      termoDoc.replaceText(
      '{AAAA}', parametros.ano
      ),
      termoDoc.replaceText(
      '{MESINGLES}', parametros.mesingles
      ),
      termoDoc.replaceText(
      '{TIPO}', parametros.tipo
      ),
      termoDoc.replaceText(
      '{DHDIA}', parametros.diahoje
      ),
      termoDoc.replaceText(
      '{DHMESINGLES}', parametros.meshoje
      ),
      termoDoc.replaceText(
      '{DHANO}', parametros.anohoje
      )
    //Aqui é substituído tudo que tiver no texto no Termo de Prorrogação, por exemplo, onde está escrito {NOME} no texto, será substituído pela variável parametros.nome.
    
    }
    }
}
