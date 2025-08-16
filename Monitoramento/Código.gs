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
 * Função para buscar lista de contato
 */
function buscaDadosOtimizada(spreadsheetId, abaNome, nomeCabecalhoColuna, valorBusca, valorParaDetectarCabecalho = null) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const aba = ss.getSheetByName(abaNome);
    if (!aba) {
      Logger.log(`A aba "${abaNome}" não foi encontrada.`);
      return null;
    }

    const linhaDeCabecalhos = 2; // Linha onde estão os cabeçalhos mesclados (ex: A, B)
    const linhaDeDadosDeVerificacao = linhaDeCabecalhos + 1; // Terceira linha da tabela
    const colunaInicioRange = 2; // Coluna B (índice 2)
    const colunaFimRange = 23;   // Coluna W (índice 23)

    // PASSO 1: LER APENAS AS LINHAS ESSENCIAIS PARA DETECTAR AS TABELAS
    const linhaCabecalhosCompleta = aba.getRange(
      linhaDeCabecalhos,
      colunaInicioRange,
      1,
      colunaFimRange - colunaInicioRange + 1
    ).getValues()[0];

    const linhaDeVerificacaoCompleta = aba.getRange(
      linhaDeDadosDeVerificacao,
      colunaInicioRange,
      1,
      colunaFimRange - colunaInicioRange + 1
    ).getValues()[0];

    let colIndexAtualNoArray = 0; // Índice 0-based no array linhaCabecalhosCompleta
    let intervalosDasTabelas = [];

    // Loop principal para detectar os intervalos das tabelas
    while (colIndexAtualNoArray < linhaCabecalhosCompleta.length) {
      let valorCabecalho = String(linhaCabecalhosCompleta[colIndexAtualNoArray]).trim();

      if (valorCabecalho !== "") { // Início de uma possível tabela
        let colunaInicioLetra = String.fromCharCode(64 + colunaInicioRange + colIndexAtualNoArray);
        let larguraTabelaAtual = 0;
        let alturaTabelaAtual = 0;

        // Determinar a largura da tabela atual baseando-se na linha de verificação
        for (let j = colIndexAtualNoArray; j < linhaDeVerificacaoCompleta.length; j++) {
          let valorCelulaNaLinhaDeVerificacao = String(linhaDeVerificacaoCompleta[j]).trim();
          if (j === colIndexAtualNoArray || valorCelulaNaLinhaDeVerificacao !== "") {
            larguraTabelaAtual++;
          } else {
            break; // Tabela termina aqui (célula vazia na linha de verificação)
          }
        }        

        const ultimaLinhaComDadosDaAba = aba.getLastRow();
        const ultimaLinhaDaTabela = ultimaLinhaComDadosDaAba; // Assumindo que a tabela vai até o fim da aba

        let letraColunaFimTabela = String.fromCharCode(64 + colunaInicioRange + colIndexAtualNoArray + larguraTabelaAtual - 1);
        if (colunaInicioRange + colIndexAtualNoArray + larguraTabelaAtual - 1 > colunaFimRange) {
          letraColunaFimTabela = String.fromCharCode(64 + colunaFimRange);
          larguraTabelaAtual = colunaFimRange - (colunaInicioRange + colIndexAtualNoArray) + 1;
        }

        let intervaloCompletoTabela = `${colunaInicioLetra}${linhaDeDadosDeVerificacao}:${letraColunaFimTabela}${ultimaLinhaDaTabela}`;
        intervalosDasTabelas.push(intervaloCompletoTabela);

        Logger.log(`Tabela ${intervalosDasTabelas.length} identificada: ${intervaloCompletoTabela}, Largura: ${larguraTabelaAtual} colunas`);

        colIndexAtualNoArray += larguraTabelaAtual; // Avança para a próxima possível coluna de cabeçalho
      } else {
        colIndexAtualNoArray++; // Célula de cabeçalho vazia, avança
      }
    }

    // PASSO 2: ITERAR PELOS INTERVALOS ENCONTRADOS E REALIZAR A BUSCA (ÚNICA LEITURA POR TABELA)
    // Variável de retorno único
    const resultadosBusca = {}; // Objeto para armazenar os resultados de cada tabela

    // Itera sobre cada intervalo de tabela encontrado
    for (let i = 0; i < intervalosDasTabelas.length; i++) {
      const intervaloTabela = intervalosDasTabelas[i];
      Logger.log(`Buscando em tabela: ${intervaloTabela}`);

      // LER OS DADOS APENAS DESTA TABELA ESPECÍFICA (UMA CHAMADA À API POR TABELA)
      const dadosDaTabelaAtual = aba.getRange(intervaloTabela).getValues();

      if (dadosDaTabelaAtual.length === 0) {
        Logger.log(`Tabela ${intervaloTabela} está vazia.`);
        continue; // Pula para a próxima tabela se estiver vazia
      }

      // Os cabeçalhos da tabela estão na primeira linha do 'dadosDaTabelaAtual'
      const headersTabelaAtual = dadosDaTabelaAtual[0];
      const colunaIndexTabelaAtual = headersTabelaAtual.findIndex(header =>
        String(header).trim().toUpperCase() === String(nomeCabecalhoColuna).trim().toUpperCase()
      );

      if (colunaIndexTabelaAtual === -1) {
        Logger.log(`Coluna com cabeçalho "${nomeCabecalhoColuna}" não encontrada na tabela ${intervaloTabela}.`);
        continue; // Pula para a próxima tabela
      }

      let linhaEncontradaNessaTabela = null;
      if (intervaloTabela.includes("R")) {
        linhaEncontradaNessaTabela = {
          sub_tab1: null,
          sub_tab2: null
        };
        let sub_tab1 = aba.getRange("R3:W5").getValues();
        Logger.log(`Dados da Sub Tabela 1 antes do For ${sub_tab1}`);
        let sub_tab2;
        if (String(aba.getRange("R7").getValues()).trim().toLowerCase() == "divisão") {
          let sub_tab2 = aba.getRange("R7:V10").getValues();
          Logger.log(`Dados da Sub Tabela 2 antes do For R7:V9 ${sub_tab2}`);
          for (let j = 1; j < sub_tab2.length; j++) {
          if (String(sub_tab2[j][colunaIndexTabelaAtual]).trim().toUpperCase() === valorBusca.toString().toUpperCase()) {
            linhaEncontradaNessaTabela.sub_tab2 = sub_tab2[j];
            break; // Encontrou a linha, pode sair do loop interno
          }
        }
        } else {
          let sub_tab2 = aba.getRange("R8:V11").getValues();
          Logger.log(`Dados da Sub Tabela 2 antes do For R8:V10 ${sub_tab2}`);
          for (let j = 1; j < sub_tab2.length; j++) {
          if (String(sub_tab2[j][colunaIndexTabelaAtual]).trim().toUpperCase() === valorBusca.toString().toUpperCase()) {
            linhaEncontradaNessaTabela.sub_tab2 = sub_tab2[j];
            break; // Encontrou a linha, pode sair do loop interno
          }
        }
        }        
        for (let j = 1; j < sub_tab1.length; j++) {
          if (String(sub_tab1[j][colunaIndexTabelaAtual]).trim().toUpperCase() === valorBusca.toString().toUpperCase()) {
            linhaEncontradaNessaTabela.sub_tab1 = sub_tab1[j];
            break; // Encontrou a linha, pode sair do loop interno
          }
        }
        Logger.log(`Dados da linha Sub1+Sub2 ${linhaEncontradaNessaTabela.sub_tab1} | ${linhaEncontradaNessaTabela.sub_tab2}`);
      } else {
        // Começa a busca a partir da SEGUNDA linha da tabela (índice 1 no array 'dadosDaTabelaAtual')
        for (let j = 1; j < dadosDaTabelaAtual.length; j++) {
          if (String(dadosDaTabelaAtual[j][colunaIndexTabelaAtual]).trim().toUpperCase() === valorBusca.toString().toUpperCase()) {
            linhaEncontradaNessaTabela = dadosDaTabelaAtual[j];
            break; // Encontrou a linha, pode sair do loop interno
          }
        }
      }

      // Armazenar o resultado no objeto de resultados
      // Você pode definir chaves mais significativas, como 'tabelaHorarioComercial', 'tabelaPosComercial', etc.
      // Ou simplesmente numerar 'tabela1', 'tabela2', 'tabela3'.
      resultadosBusca[`tabela${i + 1}`] = {
        intervalo: intervaloTabela,
        headers: headersTabelaAtual,
        data: linhaEncontradaNessaTabela
      };
    }

    // Retorna todos os resultados encontrados
    if (Object.keys(resultadosBusca).length > 0) {
        return resultadosBusca;
    } else {
        return null; // Nenhuma tabela ou linha encontrada
    }

  } catch (error) {
    Logger.log(`Erro ao Processar a busca de dados: ${error}`);
    throw new Error(`Erro ao Processar a busca de dados: ${error}`);
  }
}
