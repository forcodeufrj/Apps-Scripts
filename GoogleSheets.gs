function PegarNomes() {
  var ui = SpreadsheetApp.getUi();
  
  // Prompt para obter o nome do Membro, sua Diretoria e seu Projeto
  var response = ui.prompt('Por favor, insira o nome do membro, sua Diretoria (Projetos, Marketing, Pessoas e Presidência) e seu projeto (C, Python, VBA) separados por ponto e vírgula (;):', ui.ButtonSet.OK_CANCEL);
  
  // Verifica se o usuário clicou em "OK" e obteve a resposta
  if (response.getSelectedButton() == ui.Button.OK) {
    var input = response.getResponseText();
    
    // Separa os dados pelo ponto e vírgula
    var partes = input.split(';');
    if (partes.length != 3) {
      ui.alert('Erro', 'Por favor, insira três dados separados por ponto e vírgula.', ui.ButtonSet.OK);
      return;
    }
  
    // Divisão do que cada termo significa
    var NomeMembro = partes[0].trim();
    var DiretoriaMembro = partes[1].trim();
    var ProjetoMembro = partes[2].trim();
    
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    
    // Olhando a qual projeto e Diretoria pertence 
    //e formatando a célula de acordo com verificação 
    //se ela está vazia
    if (ProjetoMembro == "C"){
      if (DiretoriaMembro == "Projetos"){
        //pega a coluna B a partir da B7
        var alocacao = planilha.getSheets()[0];
        var colunaB = planilha.getRange('B7:B');
        var valores = colunaB.getValues();
        var linhaIndex = 0;
        for (var i = 0; i < valores.length; i++) {
          if (!valores[i][0]) {
            linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
            break
          }
        }
        // Vou preencher a primeira célula vazia da coluna
        if (linhaIndex > 0) {
          var celula = planilha.getRange('B' + linhaIndex);
          celula.setValue(NomeMembro);
          celula.setFontColor('#00ffff');
        } 
      }   

      // mesmo código, para outra Diretoria
      else if (DiretoriaMembro == "Marketing"){
        //pega a coluna B a partir da B7
        var alocacao = planilha.getSheets()[0];
        var colunaB = planilha.getRange('B7:B');
        var valores = colunaB.getValues();
        var linhaIndex = 0;
        for (var i = 0; i < valores.length; i++) {
          if (!valores[i][0]) {
            linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
            break;
          }
        }
        if (linhaIndex > 0) {
          var celula = planilha.getRange('B' + linhaIndex);
          celula.setValue(NomeMembro);
          celula.setFontColor('#00f00a');
        } 
      }   
      // mesmo código, para outra Diretoria
      else if (DiretoriaMembro == "Pessoas"){
        //pega a coluna B a partir da B7
        var alocacao = planilha.getSheets()[0];
        var colunaB = planilha.getRange('B7:B');
        var valores = colunaB.getValues();
        var linhaIndex = 0;
        for (var i = 0; i < valores.length; i++) {
          if (!valores[i][0]) {
            linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
            break;
          }
        }
        if (linhaIndex > 0) {
          var celula = planilha.getRange('B' + linhaIndex);
          celula.setValue(NomeMembro);
          celula.setFontColor('#FFFF00');
        } 
      }   
      // mesmo código, para outra Diretoria
      else if (DiretoriaMembro == "Presidência"){
        //pega a coluna B a partir da B7
        var alocacao = planilha.getSheets()[0];
        var colunaB = planilha.getRange('B7:B');
        var valores = colunaB.getValues();
        var linhaIndex = 0;
        for (var i = 0; i < valores.length; i++) {
          if (!valores[i][0]) {
            linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
            break;
          }
        }
        if (linhaIndex > 0) {
          var celula = planilha.getRange('B' + linhaIndex);
          celula.setValue(NomeMembro);
          celula.setFontColor('#FF0000');
        }
      }
      // Caso coloquem uma Diretoria que não exista
      else{
        ui.alert('Erro', 'Por favor, as Diretorias são: Projetos, Marketing, Pessoas e Presidência', ui.ButtonSet.OK);
        return;
      }
    }

    // código igual, na coluna do Python (C). 
    // O ELSE IF é importante!!!
    else if (ProjetoMembro == "Python"){
        if (DiretoriaMembro == "Projetos"){
          //pega a coluna B a partir da B7
          var alocacao = planilha.getSheets()[0];
          var colunaB = planilha.getRange('C7:C');
          var valores = colunaB.getValues();
          var linhaIndex = 0;
          for (var i = 0; i < valores.length; i++) {
            if (!valores[i][0]) {
              linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
              break
            }
          }
          if (linhaIndex > 0) {
            var celula = planilha.getRange('C' + linhaIndex);
            celula.setValue(NomeMembro);
            celula.setFontColor('#00FFFF');
          } 
        }   

        else if (DiretoriaMembro == "Marketing"){
          //pega a coluna B a partir da B7
          var alocacao = planilha.getSheets()[0];
          var colunaB = planilha.getRange('C7:C');
          var valores = colunaB.getValues();
          var linhaIndex = 0;
          for (var i = 0; i < valores.length; i++) {
            if (!valores[i][0]) {
              linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
              break;
            }
          }
          if (linhaIndex > 0) {
            var celula = planilha.getRange('C' + linhaIndex);
            celula.setValue(NomeMembro);
            celula.setFontColor('#00f00a');
          } 
        }   

        else if (DiretoriaMembro == "Pessoas"){
          //pega a coluna B a partir da B7
          var alocacao = planilha.getSheets()[0];
          var colunaB = planilha.getRange('C7:C');
          var valores = colunaB.getValues();
          var linhaIndex = 0;
          for (var i = 0; i < valores.length; i++) {
            if (!valores[i][0]) {
              linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
              break;
            }
          }
          if (linhaIndex > 0) {
            var celula = planilha.getRange('C' + linhaIndex);
            celula.setValue(NomeMembro);
            celula.setFontColor('#FFFF00');
          } 
        }   

        else if (DiretoriaMembro == "Presidência"){
          //pega a coluna B a partir da B7
          var alocacao = planilha.getSheets()[0];
          var colunaB = planilha.getRange('C7:C');
          var valores = colunaB.getValues();
          var linhaIndex = 0;
          for (var i = 0; i < valores.length; i++) {
            if (!valores[i][0]) {
              linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
              break;
            }
          }
          if (linhaIndex > 0) {
            var celula = planilha.getRange('C' + linhaIndex);
            celula.setValue(NomeMembro);
            celula.setFontColor('#FF0000');
          }
        }else{
          ui.alert('Erro', 'Por favor, as Diretorias são: Projetos, Marketing, Pessoas e Presidência', ui.ButtonSet.OK);
          return;
        }
    }
    
    // código igual, na coluna do VBA (D)
    else if (ProjetoMembro == "VBA") {
        if (DiretoriaMembro == "Projetos") {
          //pega a coluna B a partir da B7
          var alocacao = planilha.getSheets()[0];
          var colunaB = planilha.getRange('D7:D');
          var valores = colunaB.getValues();
          var linhaIndex = 0;
          for (var i = 0; i < valores.length; i++) {
            if (!valores[i][0]) {
              linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
              break
            }
          }
          if (linhaIndex > 0) {
            var celula = planilha.getRange('D' + linhaIndex);
            celula.setValue(NomeMembro);
            celula.setFontColor('#00FFFF');
          } 
        }   

        else if (DiretoriaMembro == "Marketing"){
          //pega a coluna B a partir da B7
          var alocacao = planilha.getSheets()[0];
          var colunaB = planilha.getRange('D7:D');
          var valores = colunaB.getValues();
          var linhaIndex = 0;
          for (var i = 0; i < valores.length; i++) {
            if (!valores[i][0]) {
              linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
              break;
            }
          }
          if (linhaIndex > 0) {
            var celula = planilha.getRange('D' + linhaIndex);
            celula.setValue(NomeMembro);
            celula.setFontColor('#00f00a');
          } 
        }   

        else if (DiretoriaMembro == "Pessoas"){
          //pega a coluna B a partir da B7
          var alocacao = planilha.getSheets()[0];
          var colunaB = planilha.getRange('D7:D');
          var valores = colunaB.getValues();
          var linhaIndex = 0;
          for (var i = 0; i < valores.length; i++) {
            if (!valores[i][0]) {
              linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
              break;
            }
          }
          if (linhaIndex > 0) {
            var celula = planilha.getRange('D' + linhaIndex);
            celula.setValue(NomeMembro);
            celula.setFontColor('#FFFF00');
          } 
        }   

        else if (DiretoriaMembro == "Presidência"){
          //pega a coluna B a partir da B7
          var alocacao = planilha.getSheets()[0];
          var colunaB = planilha.getRange('D7:D');
          var valores = colunaB.getValues();
          var linhaIndex = 0;
          for (var i = 0; i < valores.length; i++) {
            if (!valores[i][0]) {
              linhaIndex = i + 7; // Adiciona 7 para ajustar o índice da linha
              break;
            }
          }
          if (linhaIndex > 0) {
            var celula = planilha.getRange('D' + linhaIndex);
            celula.setValue(NomeMembro);
            celula.setFontColor('#FF0000');
          } 
        }else{
          ui.alert('Erro', 'Por favor, as Diretorias são: Projetos, Marketing, Pessoas e Presidência', ui.ButtonSet.OK);
          return;
        }  
      }
    // Caso coloque um Projeto não disponível na Liga   
    else{
      ui.alert('Erro', 'Por favor, os projetos disponíveis são: VBA, Python e C.', ui.ButtonSet.OK);
      return;
    }
    var cor = '#614ad3'; // Cor da Liga
    celula.setBorder(true, true, true, true, true, true, cor, SpreadsheetApp.BorderStyle.SOLID_THICK); 
  } 
}