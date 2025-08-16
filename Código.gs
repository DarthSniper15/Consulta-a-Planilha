/**
 * Chama o front-end
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}

/**
 * Pega as abas na planilha selecionada
 */
function getNomesAbas(spreadsheetId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const abas = ss.getSheets().map(sheet => sheet.getName());
    return abas;
  } catch (error) {
    Logger.log(`Erro ao obter nomes das abas: ${error}`);
    throw new Error(`Erro ao obter nomes das abas: ${error}`); // Lança o erro para ser capturado no frontend
  }
}

/**
 * Busca a coluna na aba informada
 */
function buscarColuna(spreadsheetId, nomeAba, colunaBusca) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      Logger.log(`Aba "${nomeAba}" não encontrada.`);
      throw new Error(`Aba "${nomeAba}" não encontrada.`);
    }

    let colunaIndex;
    if (colunaBusca.length === 1 && colunaBusca.match(/[A-Z]/i)) {
      colunaIndex = colunaBusca.toUpperCase().charCodeAt(0) - 65 + 1; // Apps Script usa índice baseado em 1 para getRange
    } else {
      // Se a entrada não for uma letra única, tenta encontrar pelo nome do cabeçalho (primeira linha)
      const primeiraLinha = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
      colunaIndex = primeiraLinha.findIndex(header => header.toUpperCase() === colunaBusca.toUpperCase()) + 1;
      if (colunaIndex === 0) {
        Logger.log(`Coluna "${colunaBusca}" não encontrada na aba "${nomeAba}".`);
        throw new Error(`Coluna "${colunaBusca}" não encontrada na aba "${nomeAba}".`);
      }
    }

    // Obtém todos os valores da coluna especificada, desde a primeira linha até a última
    const ultimaLinha = aba.getLastRow();
    const valoresColuna = aba.getRange(1, colunaIndex, ultimaLinha).getValues().flat().filter(String); // Pega os valores, transforma em array plano e remove vazios

    return valoresColuna;

  } catch (error) {
    Logger.log(`Erro ao buscar coluna: ${error}`);
    throw new Error(`Erro ao buscar coluna: ${error}`);
  }
}

/**
 * Busca as informações na linha da TSS selecionada
 */
function buscarLinha(spreadsheetId, nomeAba, nomeCabecalhoColuna, valorBusca, valorParaDetectarCabecalho = null) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      return null;
    }

    const dados = aba.getDataRange().getValues();
    if (dados.length === 0) {
      return null;
    }

    let linhaCabecalho = 0;

    // Busca a primeira linha que contém o valor pré-definido (aplicando trim na comparação)
    if (valorParaDetectarCabecalho) {
      const termoBuscaCabecalho = String(valorParaDetectarCabecalho).trim().toUpperCase();
      for (let i = 0; i < dados.length; i++) {
        if (dados[i].some(cell => String(cell).trim().toUpperCase() === termoBuscaCabecalho)) {
          linhaCabecalho = i + 1;
          break;
        }
      }
      if (linhaCabecalho === 0) {
        Logger.log(`Não foi possível encontrar a linha de cabeçalho com o valor "${valorParaDetectarCabecalho}".`);
        return null;
      }
    } else {
      linhaCabecalho = 1; // Se nenhum valor para detectar cabeçalho for fornecido, assume a primeira linha
    }

    const headers = dados[linhaCabecalho - 1];
    const colunaIndex = headers.findIndex(header => String(header).trim().toUpperCase() === String(nomeCabecalhoColuna).trim().toUpperCase());

    if (colunaIndex === -1) {
      Logger.log(`Coluna com cabeçalho "${nomeCabecalhoColuna}" não encontrada na linha ${linhaCabecalho}.`);
      return null;
    }

    let linhaEncontrada = null;
    for (let i = linhaCabecalho; i < dados.length; i++) {
      // Aplica trim() ao valor da célula antes da comparação
      if (String(dados[i][colunaIndex]).trim().toUpperCase() === valorBusca.toString().toUpperCase()) {
        linhaEncontrada = dados[i];
        break;
      }
    }

    if (linhaEncontrada) {
      return { headers: headers, data: linhaEncontrada };
    } else {
      return null;
    }

  } catch (error) {
    Logger.log(`Erro ao buscar dados: ${error}`);
    throw new Error(`Erro ao buscar dados: ${error}`);
  }
}

/**
 * Função teste
 * Funções para pegar as orientações especiais para cada TSS especial
 */
function osEspeciais(spreadsheetId, tss, nomeAba){
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      return null;
    }

    const dados = aba.getDataRange().getValues();
    if (dados.length === 0) {
      return null;
    }

    let linhaCabecalho = 0;

    // Busca a primeira linha que contém o valor pré-definido (aplicando trim na comparação)
    if (valorParaDetectarCabecalho) {
      const termoBuscaCabecalho = String(valorParaDetectarCabecalho).trim().toUpperCase();
      for (let i = 0; i < dados.length; i++) {
        if (dados[i].some(cell => String(cell).trim().toUpperCase() === termoBuscaCabecalho)) {
          linhaCabecalho = i + 1;
          break;
        }
      }
      if (linhaCabecalho === 0) {
        Logger.log(`Não foi possível encontrar a linha de cabeçalho com o valor "${valorParaDetectarCabecalho}".`);
        return null;
      }
    } else {
      linhaCabecalho = 1; // Se nenhum valor para detectar cabeçalho for fornecido, assume a primeira linha
    }

    const headers = dados[linhaCabecalho - 1];
    const colunaIndex = headers.findIndex(header => String(header).trim().toUpperCase() === String(nomeCabecalhoColuna).trim().toUpperCase());

    if (colunaIndex === -1) {
      Logger.log(`Coluna com cabeçalho "${nomeCabecalhoColuna}" não encontrada na linha ${linhaCabecalho}.`);
      return null;
    }

    let linhaEncontrada = null;
    for (let i = linhaCabecalho; i < dados.length; i++) {
      // Aplica trim() ao valor da célula antes da comparação
      if (String(dados[i][colunaIndex]).trim().toUpperCase() === valorBusca.toString().toUpperCase()) {
        linhaEncontrada = dados[i];
        break;
      }
    }

    if (linhaEncontrada) {
      return { headers: headers, data: linhaEncontrada };
    } else {
      return null;
    }

  } catch (error) {
    Logger.log(`Erro ao buscar dados: ${error}`);
    throw new Error(`Erro ao buscar dados: ${error}`);
  }
}

/**
 * Pega Vazamentos e Arrebentado
 */
function osEspeciaisVazamento(spreadsheetId, abaSelecionada, tss, resultado){
  try {    
    // Código para divisão OJMI
    if (spreadsheetId == "id_planilha"){
      // Pega a aba de vazamento e arrebentado
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const aba = ss.getSheetByName("Vazamentos e arrebentados");
      if (aba) {
        Logger.log(`Aba selecionada: ${aba.getName()}`); // Loga o nome da aba
      } else {
        Logger.log(`Aba "Vazamentos e arrebentados" não encontrada.`); // Loga se a aba não foi encontrada
      }

      Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)


      let dados_info = aba.getRange("C2:H7").getValues();
      let dados_exec = aba.getRange("C9:F14").getValues();
      let x = 0;

      Logger.log(`Informantes ${dados_info} Executantes ${dados_exec}`);

      if(abaSelecionada.trim().toLowerCase() == "cabreuva"){
        abaSelecionada = "cabreúva";
      }
      if (tss.trim().toLowerCase() == "vazamento nao visivel cavalete"){
        while (x <= 5){
          if (String(dados_info[x][0]).trim().toLowerCase() == abaSelecionada.trim().toLowerCase()){
            Logger.log(`Linha Selecionada na cidade ${dados_info[x][0]}`);
            return {info1: dados_info[x][2], info2: dados_info[x][3], info3: dados_info[x][4], exec: dados_exec[x][2], resultado: resultado, tss: "vazamento"};
            break;
          } else {
            x++;
            Logger.log(`Não encontrado, indo para o próximo índice ${x} Célula escaneada ${dados_info[x][0].toLowerCase()} Valor Procurado ${abaSelecionada}`);
          }
        }
      } else {
        while (x <= 5){
          if (String(dados_info[x][0]).trim().toLowerCase() == abaSelecionada.trim().toLowerCase()){
            Logger.log(`Linha Selecionada na cidade ${dados_info[x][0]}`);
            return {info1: dados_info[x][2], info2: dados_info[x][3], info3: dados_info[x][4], exec: dados_exec[x][3], resultado: resultado, tss: "vazamento"};
            break;
          } else {
            x++;
            Logger.log(`Não encontrado, indo para o próximo índice ${x} Célula escaneada ${dados_info[x][0].toLowerCase()} Valor Procurado ${abaSelecionada}`);
          }
        }
      }
    //Código para divisão OJMC
    } else if (spreadsheetId == "id_planilha"){
      // Pega a aba de vazamento e arrebentado
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const vpp = ss.getSheetByName("Vazamentos e arrebentados");
      const clpp = ss.getSheetByName("Vazamentos e arrebentados (2)");
      /*if (aba) {
        Logger.log(`Aba selecionada: ${aba.getName()}`); // Loga o nome da aba
      } else {
        Logger.log(`Aba "Vazamentos e arrebentados" não encontrada.`); // Loga se a aba não foi encontrada
      }*/

      let dadosVP = vpp.getRange("C4:F4").getValues();
      let dadosCLP = clpp.getRange("C4:E4").getValues();

      Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

      if (abaSelecionada == "vp") {
        if (!tss.trim().toLowerCase().includes("rede")) {
          return {exec: dadosVP[0][2], resultado: resultado, cida: abaSelecionada, tss: "vazamento"};
        } else {
          return {exec: dadosVP[0][1], resultado: resultado, cida: abaSelecionada, tss: "vazamento"};
        }
      } else if (abaSelecionada == "clp") {
        return {cida: abaSelecionada, resultado: resultado, tss: "vazamento"};
      } else {
        throw new Error(`Erro ao buscar aba vazamento | Cidade não localizada`);
      }
    } else {
      throw new Error(`Erro ao buscar vazamento | Unidade não encontrada ou Inexistente`);
    }

    

  } catch (error) {
    Logger.log(`Erro ao buscar dados: ${error}`);
    throw new Error(`Erro ao buscar dados: ${error}`);
  }
}

/**
 * Busca orientações de falta de água
 */
function osFaltaAgua(spreadsheetId, abaSelecionada, tss, resultado){
  try {    
    // Código para divisão OJMI
    if (spreadsheetId == "id_planilha"){
      // Pega a aba de vazamento e arrebentado
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const aba = ss.getSheetByName("Falta de água");
      if (aba) {
        Logger.log(`Aba selecionada: ${aba.getName()}`); // Loga o nome da aba
      } else {
        Logger.log(`Aba "Falta de água" não encontrada.`); // Loga se a aba não foi encontrada
      }

      Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

      let dados_exec = aba.getRange("C7:F11").getValues();
      let afas_feri = aba.getRange("E7:E11").getValues();
      let subst = aba.getRange("F7:F11").getValues();
      let x = 0;

      Logger.log(`Executantes ${dados_exec}`);

      if(abaSelecionada.trim().toLowerCase() == "cabreuva"){
        abaSelecionada = "cabreúva";
      }
      while (x <= 5){
        if (String(dados_exec[x][0]).trim().toLowerCase() == abaSelecionada.trim().toLowerCase()){
          Logger.log(`Linha Selecionada na cidade ${dados_exec[x][0]}`);
          Logger.log(`Os seguintes dados serão enviados para o Front-End: ${afas_feri[x][0]} | ${subst[x]}`);
          return {exec: dados_exec[x][1], afasferi: afas_feri[x][0], subst: subst[x], resultado: resultado, tss: "falta"};
          break;
        } else {
          x++;
          Logger.log(`Não encontrado, indo para o próximo índice ${x} Célula escaneada ${dados_exec[x][0].toLowerCase()} Valor Procurado ${abaSelecionada}`);
        }
      }
    //Código para divisão OJMC
    } else if (spreadsheetId == "id_planilha"){
      // Pega a aba de Falta de água
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const vpp = ss.getSheetByName("Falta de água");

      let dados_exec = vpp.getRange("C7:F8").getValues();
      let afas_feri = vpp.getRange("E7:E8").getValues();
      let subst = vpp.getRange("F7:F8").getValues();

      Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

      if (abaSelecionada == "vp") {
        return {exec: dados_exec[0][1], afasferi: afas_feri[0][0], subst: subst[0][0], resultado: resultado, tss: "falta"};
      } else if (abaSelecionada == "clp") {
        return {exec: dados_exec[1][1], afasferi: afas_feri[1][0], subst: subst[1][0], resultado: resultado, tss: "falta"};
      } else {
        throw new Error(`Erro ao buscar aba vazamento | Cidade não localizada`);
      }
    // Código para divisão OJMH   
    } else if (spreadsheetId == "id_planilha") {   
      // Pega a aba de vazamento e arrebentado
      const ss = SpreadsheetApp.openById(spreadsheetId);
      const aba = ss.getSheetByName("Falta de água ");
      if (aba) {
        Logger.log(`Aba selecionada: ${aba.getName()}`); // Loga o nome da aba
      } else {
        Logger.log(`Aba "Falta de água" não encontrada.`); // Loga se a aba não foi encontrada
      }

      Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

      let dados_exec = aba.getRange("C7:F12").getValues();
      let afas_feri = aba.getRange("E7:E12").getValues();
      let subst = aba.getRange("F7:F12").getValues();
      let x = 0;

      Logger.log(`Executantes ${dados_exec}`);

      /*if(abaSelecionada.trim().toLowerCase() == "hortolandia"){
        abaSelecionada = "hortolândia";
      }*/
      while (x <= 6){
        if (String(dados_exec[x][0]).trim().toLowerCase() == abaSelecionada.trim().toLowerCase()){
          Logger.log(`Linha Selecionada na cidade ${dados_exec[x][0]}`);
          Logger.log(`Os seguintes dados serão enviados para o Front-End: ${afas_feri[x][0]} | ${subst[x]}`);
          return {exec: dados_exec[x][1], afasferi: afas_feri[x][0], subst: subst[x], resultado: resultado, tss: "falta"};
          break;
        } else {
          x++;
          Logger.log(`Não encontrado, indo para o próximo índice ${x} Célula escaneada ${dados_exec[x][0].toLowerCase()} Valor Procurado ${abaSelecionada}`);
        }
      }
    } else {
      throw new Error(`Erro ao buscar vazamento | Unidade não encontrada ou Inexistente`);
    }

  } catch (error) {
    Logger.log(`Erro ao buscar dados: ${error}`);
    throw new Error(`Erro ao buscar dados: ${error}`);
  }
}

/**
 * Busca orientações de Ordens especiais
 */
function osEspeciaisGeral(spreadsheetId, abaSelecionada, tss, resultado){
  try {    
    // Código para Serviço Solicitado
    Logger.log(`Ordem a ser localizado ${tss}`);
    if (tss.includes("solicitado") || tss.includes("responsabilidade") || tss.includes("executado")){
      // Código para divisão OJMI
      if (spreadsheetId == "id_planilha"){
        // Pega a aba de vazamento e arrebentado
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const aba = ss.getSheetByName("Solicitados e executados");
        if (aba) {
          Logger.log(`Aba selecionada: ${aba.getName()}`); // Loga o nome da aba
        } else {
          Logger.log(`Aba "Falta de água" não encontrada.`); // Loga se a aba não foi encontrada
        }

        Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

        let dados_exec = aba.getRange("C2:H3").getValues();
        let afas_feri = aba.getRange("G2:G3").getValues();
        let subst = aba.getRange("H2:H3").getValues();
        let x = 0;
        let y = 0;
        let col_esg = "";
        let op_esg = "";
        let col_agua = "";
        let op_agua = "";

        Logger.log(`Executantes ${dados_exec}`);

        /*if(abaSelecionada.trim().toLowerCase() == "cabreuva"){
          abaSelecionada = "cabreúva";
        }*/
        /*while (x <= 5){
          if (String(dados_exec[x][0]).trim().toLowerCase() == abaSelecionada.trim().toLowerCase()){
            Logger.log(`Linha Selecionada na cidade ${dados_exec[x][0]}`);
            Logger.log(`Os seguintes dados serão enviados para o Front-End: ${afas_feri[x][0]} | ${subst[x]}`);
            return {exec: dados_exec[x][1], afasferi: afas_feri[x][0], subst: subst[x], resultado: resultado, tss: "solicitado"};
            break;
          } else {
            x++;
            Logger.log(`Não encontrado, indo para o próximo índice ${x} Célula escaneada ${dados_exec[x][0].toLowerCase()} Valor Procurado ${abaSelecionada}`);
          }
        }*/
        while (y <= 4){
          if (dados_exec[0][y].trim().toLowerCase().includes("esgoto")){
            Logger.log(`Dados escaneados - Esgoto: ${dados_exec[0][y]}`);
            //col_esg = dados_exec[0];
            op_esg = dados_exec[1][1];
            y++;
            Logger.log(`Dados enviados - Esgoto: ${dados_exec[1][1]}`);
          } else if (dados_exec[0][y].trim().toLowerCase().includes("água")){
            Logger.log(`Dados escaneados - Água: ${dados_exec[0][y]}`);
            //col_agua = dados_exec[0];
            op_agua = dados_exec[1][2];
            y++;
            Logger.log(`Dados enviados - Água: ${dados_exec[1][2]}`);
          } else if (op_esg != "" && op_agua != ""){
            return {exec_agu: op_agua, exec_esg: op_esg, afasferi: afas_feri[1], subst: subst[1], resultado: resultado, tss: "solicitado"};
            break;
          } else {
            Logger.log(`Não encontrado, indo para o próximo índice ${y} Célula escaneada ${dados_exec[y][0].toLowerCase()}`);
            y++;
          }          
          //Logger.log(`Serão enviado o ${subst[1]} e ${afas_feri[1]} para o front-end`);
        }
      //Código para divisão OJMC
      } else if (spreadsheetId == "id_planilha"){
        // Pega a aba de Falta de água
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const vpp = ss.getSheetByName("Solicitados e executados");

        let dados_exec = vpp.getRange("C7:F8").getValues();
        let afas_feri = vpp.getRange("E7:E8").getValues();
        let subst = vpp.getRange("F7:F8").getValues();

        Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

        if (abaSelecionada == "vp") {
          return {exec: dados_exec[0][1], afasferi: afas_feri[0][0], subst: subst[0][0], resultado: resultado, tss: "falta"};
        } else if (abaSelecionada == "clp") {
          return {exec: dados_exec[1][1], afasferi: afas_feri[1][0], subst: subst[1][0], resultado: resultado, tss: "falta"};
        } else {
          throw new Error(`Erro ao buscar aba vazamento | Cidade não localizada`);
        }
      // Código para divisão OJMH   
      } else if (spreadsheetId == "id_planilha") {   
        // Pega a aba de vazamento e arrebentado
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const aba = ss.getSheetByName("Solicitados e executados");
        if (aba) {
          Logger.log(`Aba selecionada: ${aba.getName()}`); // Loga o nome da aba
        } else {
          Logger.log(`Aba "Falta de água" não encontrada.`); // Loga se a aba não foi encontrada
        }

        Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

        let dados_exec = aba.getRange("C7:F12").getValues();
        let afas_feri = aba.getRange("E7:E12").getValues();
        let subst = aba.getRange("F7:F12").getValues();
        let x = 0;

        Logger.log(`Executantes ${dados_exec}`);

        /*if(abaSelecionada.trim().toLowerCase() == "hortolandia"){
          abaSelecionada = "hortolândia";
        }*/
        while (x <= 6){
          if (String(dados_exec[x][0]).trim().toLowerCase() == abaSelecionada.trim().toLowerCase()){
            Logger.log(`Linha Selecionada na cidade ${dados_exec[x][0]}`);
            Logger.log(`Os seguintes dados serão enviados para o Front-End: ${afas_feri[x][0]} | ${subst[x]}`);
            return {exec: dados_exec[x][1], afasferi: afas_feri[x][0], subst: subst[x], resultado: resultado, tss: "falta"};
            break;
          } else {
            x++;
            Logger.log(`Não encontrado, indo para o próximo índice ${x} Célula escaneada ${dados_exec[x][0].toLowerCase()} Valor Procurado ${abaSelecionada}`);
          }
        }
      } else {
        throw new Error(`Erro ao buscar vazamento | Unidade não encontrada ou Inexistente`);
      }

      // Código para Religação | Restabelecer
    } else if (tss.includes("religar") || tss.includes("restabelecer")) {
      // Código para divisão OJMI
      if (spreadsheetId == "id_planilha"){
        // Pega a aba de vazamento e arrebentado
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const aba = ss.getSheetByName("Resta e religar OP");
        if (aba) {
          Logger.log(`Aba selecionada: ${aba.getName()}`); // Loga o nome da aba
        } else {
          Logger.log(`Aba "Falta de água" não encontrada.`); // Loga se a aba não foi encontrada
        }

        Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

        let dados_exec = aba.getRange("C2:D7").getValues();
        /*let afas_feri = aba.getRange("E7:E11").getValues();
        let subst = aba.getRange("F7:F11").getValues();*/
        let x = 0;

        Logger.log(`Executantes ${dados_exec}`);

        if(abaSelecionada.trim().toLowerCase() == "cabreuva"){
          abaSelecionada = "cabreúva";
        }
        while (x <= 5){
          if (String(dados_exec[x][0]).trim().toLowerCase() == abaSelecionada.trim().toLowerCase()){
            Logger.log(`Linha Selecionada na cidade ${dados_exec[x][0]}`);
            //Logger.log(`Os seguintes dados serão enviados para o Front-End: ${afas_feri[x][0]} | ${subst[x]}`);
            return {exec: dados_exec[x][1], /*afasferi: afas_feri[x][0], subst: subst[x],*/ resultado: resultado, tss: "religar"};
            break;
          } else {
            x++;
            Logger.log(`Não encontrado, indo para o próximo índice ${x} Célula escaneada ${dados_exec[x][0].toLowerCase()} Valor Procurado ${abaSelecionada}`);
          }
        }
      //Código para divisão OJMC
      } else if (spreadsheetId == "id_planilha"){
        // Pega a aba de Falta de água
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const vpp = ss.getSheetByName("Resta e religar OP");

        let dados_exec = vpp.getRange("C7:F8").getValues();
        let afas_feri = vpp.getRange("E7:E8").getValues();
        let subst = vpp.getRange("F7:F8").getValues();

        Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

        if (abaSelecionada == "vp") {
          return {exec: dados_exec[0][1], afasferi: afas_feri[0][0], subst: subst[0][0], resultado: resultado, tss: "falta"};
        } else if (abaSelecionada == "clp") {
          return {exec: dados_exec[1][1], afasferi: afas_feri[1][0], subst: subst[1][0], resultado: resultado, tss: "falta"};
        } else {
          throw new Error(`Erro ao buscar aba vazamento | Cidade não localizada`);
        }
      // Código para divisão OJMH   
      } else if (spreadsheetId == "id_planilha") {   
        // Pega a aba de vazamento e arrebentado
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const aba = ss.getSheetByName("Resta e religar OP");
        if (aba) {
          Logger.log(`Aba selecionada: ${aba.getName()}`); // Loga o nome da aba
        } else {
          Logger.log(`Aba "Falta de água" não encontrada.`); // Loga se a aba não foi encontrada
        }

        Logger.log(`Iniciando tentativa de pegar a cidade mencionada: ${abaSelecionada}`)

        let dados_exec = aba.getRange("C7:F12").getValues();
        let afas_feri = aba.getRange("E7:E12").getValues();
        let subst = aba.getRange("F7:F12").getValues();
        let x = 0;

        Logger.log(`Executantes ${dados_exec}`);

        /*if(abaSelecionada.trim().toLowerCase() == "hortolandia"){
          abaSelecionada = "hortolândia";
        }*/
        while (x <= 6){
          if (String(dados_exec[x][0]).trim().toLowerCase() == abaSelecionada.trim().toLowerCase()){
            Logger.log(`Linha Selecionada na cidade ${dados_exec[x][0]}`);
            Logger.log(`Os seguintes dados serão enviados para o Front-End: ${afas_feri[x][0]} | ${subst[x]}`);
            return {exec: dados_exec[x][1], afasferi: afas_feri[x][0], subst: subst[x], resultado: resultado, tss: "falta"};
            break;
          } else {
            x++;
            Logger.log(`Não encontrado, indo para o próximo índice ${x} Célula escaneada ${dados_exec[x][0].toLowerCase()} Valor Procurado ${abaSelecionada}`);
          }
        }
      } else {
        throw new Error(`Erro ao buscar vazamento | Unidade não encontrada ou Inexistente`);
      }
    }

  } catch (error) {
    Logger.log(`Erro ao buscar dados: ${error}`);
    throw new Error(`Erro ao buscar dados: ${error}`);
  }
}
