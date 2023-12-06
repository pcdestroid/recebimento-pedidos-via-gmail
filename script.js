// Google App Script

// Procura no email com o relatorio de solicitação que estão em aberto, baixa o relatório, salva no drive e extrair as informações da planilha .xlsx para planilha do google .gs

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  let menu = ui.createMenu('Pedidos');
  menu.addItem('Carregar', 'novosPedidos');
  menu.addToUi();
}

function novosPedidos() {

  const assuntoAProcurar = 'SOLICITAÇÕES NÃO COTADAS';
  const minhaPasta = DriveApp.getFolderById("");

  // Procurar apenas uma thread com o assunto específico.
  const threads = GmailApp.search(`subject:"${assuntoAProcurar}"`, 0, 1);

  if (threads.length > 0 && threads[0].isUnread()) { // Se houver pelo menos uma thread encontrada e não foi lido.

    threads[0].markRead()

    const messages = threads[0].getMessages();
    const attachments = messages[messages.length - 1].getAttachments(); // Obter anexos da última mensagem da thread.
    for (let j = 0; j < attachments.length; j++) {
      const anexo = attachments[j];
      // remover o arquivo existente, se houver
      const arquivosAntigos = minhaPasta.getFilesByName('RELATORIO_108.XLSX');
      while (arquivosAntigos.hasNext()) {
        const arquivoAntigo = arquivosAntigos.next();
        arquivoAntigo.setTrashed(true);
      }
      // salvar o novo anexo na pasta especificada
      var arquivo = minhaPasta.createFile(anexo).setName('RELATORIO_108.XLSX');

      // Converta o arquivo para o formato Google Sheets
      var arquivoConvertido = Drive.Files.copy({}, arquivo.getId(), { convert: true });

      // Abra a planilha convertida
      var planilha = SpreadsheetApp.openById(arquivoConvertido.id);

      // Acesse a primeira guia (planilha ativa)
      var guia = planilha.getSheets()[0];

      // Obtenha os dados da planilha
      var dados = guia.getDataRange().getValues();

      //Planilha de recebimento de pedidos
      var spreadsheetId = '';
      var planilhaRecebimento = SpreadsheetApp.openById(spreadsheetId);
      var guiaRecebimento = planilhaRecebimento.getSheetByName('Pedidos');
      var dadosRecebimento = guiaRecebimento.getDataRange().getValues();

      // Verifique cada linha na planilha do anexo
      for (let l = 1; l < dados.length; l++) {
        var empresa = dados[l][1]; // Coluna B

        var solicitacao = dados[l][2]; // Coluna C

        let sequencia = dados[l][4];

        var complementoProduto = dados[l][6]; // Coluna F

        var nomeUsuario = dados[l][10]; // Coluna J

        if (solicitacao !== "" && solicitacao !== "Solicitação") {

          // Verifique se a solicitação não existe na planilha "Recebimento de Pedidos"
          var existeSolicitacao = false;
          for (var m = 0; m < dadosRecebimento.length; m++) {
            if (dadosRecebimento[m][2] === solicitacao && dadosRecebimento[m][3] === sequencia && dadosRecebimento[m][1] === empresa) { // Coluna 3
              existeSolicitacao = true;
              break;
            }
          }

          // Se a solicitação não existe, insira os dados na planilha "Recebimento de Pedidos"
          if (!existeSolicitacao) {
            guiaRecebimento.appendRow([new Date(), empresa, solicitacao, sequencia, complementoProduto, nomeUsuario, '', "Recebido"]);
          }
        }

      }

      Logger.log('Relatório atualizado')
      // Feche a planilha convertida
      DriveApp.getFileById(arquivoConvertido.id).setTrashed(true);

    }

  } else {
    Logger.log('Nenhum relatório encontrado no email')
    return false;
  }

}
