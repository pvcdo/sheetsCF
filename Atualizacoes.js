function atualizarDadosContratacaoAPI() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  var ultima_linha_dados_contratacao = planilha.getSheetByName("Dados de contratação").getRange("G2").getValue()
  
  var lin = 2
  var ticket_sem_info = planilha.getSheetByName("Tickets sem info contratacao").getRange(lin,1).getValue()
  var ticket_sem_info_destino = planilha.getSheetByName("Tickets sem info contratacao").getRange(lin,3).getValue()

  var tickets = []
  var linhas_destino = []

  while (ticket_sem_info !== ""){
    tickets.push(ticket_sem_info)
    linhas_destino.push(ticket_sem_info_destino)
    
    lin++
    ticket_sem_info = planilha.getSheetByName("Tickets sem info contratacao").getRange(lin,1).getValue()
    ticket_sem_info_destino = planilha.getSheetByName("Tickets sem info contratacao").getRange(lin,3).getValue()
  }

  if(tickets.length == 0){
    console.log("Não há tickets para serem verificados")
    return
  }

  console.log("Busca iniciando com " + tickets.length + " tickets")

  const fontes = [
    [//abc
      {
        "colunas": [0, 9, 63, 0], // Converter arrays para string JSON
        "colunas_data": [63],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "GERAL",
        "fonte": "ABC"
      },
      "https://script.google.com/macros/s/AKfycbyCDiVFryD-0uz4VDaQ5QpbdNPTGlm8bAgVKdNKr935A5hvMJ-W1evShr8GWmDPMvfK/exec"
    ],
    [//ocupacao cadm
      {
        "colunas": [32, 9, 63, 0], // Converter arrays para string JSON
        "colunas_data": [63],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 33,
        "nome_aba_fonte": "GERAL",
        "fonte": "CADM OCUPADO POR ABC"
      },
      "https://script.google.com/macros/s/AKfycbyCDiVFryD-0uz4VDaQ5QpbdNPTGlm8bAgVKdNKr935A5hvMJ-W1evShr8GWmDPMvfK/exec"
    ],
    [//agendas
      {
        "colunas": [0, 1, 5, 0], // Converter arrays para string JSON
        "colunas_data": [5],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CONTRATADOS NA AGENDA",
        "fonte": "CONTRATADOS NAS AGENDAS"
      },
      "https://script.google.com/macros/s/AKfycbxoEJgUCrHbErS9kq-aBsF46etNKJyEZyhNZCWQOw2M3Gpjd0mHRVWc1BuOTcuEgh06/exec"
    ],
    [//cadm arq
      {
        "colunas": [0, 38, 43, 0], // PROTOCOLO	MATRICULA	DATA INICIO
        "colunas_data": [43],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CONTRATAÇÕES ARQUIVADAS",
        "fonte": "CADM HISTORICO"
      },
      "https://script.google.com/macros/s/AKfycbzLpqp0v6b2mDGNTyD85XiBIBOUnEGFFcwIDNh8uzEec92sjHiT0vqpE1GUm6YK7FuZ/exec"
    ],
    [//cadm
      {
        "colunas": [0, 38, 43, 0], // PROTOCOLO	MATRICULA	DATA INICIO
        "colunas_data": [43],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CDC",
        "fonte": "CDC"
      },
      "https://script.google.com/macros/s/AKfycbx1Y2QwYCpzkOpXaGqYSi2DLO7462r7mUfjTrTsxFlyhkfw61Glhxh7uqcfwEFXJQGcDw/exec"
    ],
  ]

  var csv = ""
  
  fontes.forEach((busca) => {
      var options = {
        'method' : 'post', // Define o método como POST
        'contentType': 'application/json', // Informa que o corpo é JSON
        'payload' : JSON.stringify(busca[0]) // Converte o objeto para string JSON
      };

      var resposta = UrlFetchApp.fetch(busca[1],options)

      var texto_resposta = resposta.getContentText()

      if(texto_resposta !== ""){
        csv += texto_resposta + '\n'
      }

      console.log("finalizado " + busca[0].fonte)
    }
  )

  var dados = Utilities.parseCsv(csv);

  const tickets_dados = dados.map(dado => dado[0])

  var tickets_fora_controles = [] 
  tickets.forEach(ticket => {
    const id_ticket = tickets_dados.indexOf(ticket)
    const consta = id_ticket >= 0
    if(ticket == '31.00462193/2025-76'){
      console.log("este está fora")
    }
    if(!consta){
      tickets_fora_controles.push(ticket)
    }
  })
  
  var aba = planilha.getSheetByName('Dados de contratação');

  console.log('iniciando inclusão dos dados das consultas')

  dados.forEach(dado => {
    
    const i = tickets.indexOf(dado[0])

    if(dado[1]=="" && dado[2]=="") return

    /*if(linhas_destino[i] < 0){
      console.log("aqui é um dado que ainda não está no controle")
    }*/
    
    const linha = linhas_destino[i] < 0 ? ++ultima_linha_dados_contratacao : linhas_destino[i]
    //console.log(`linha ${i}`)
    aba.getRange(linha,1,1,dado.length).setValues([dado])
  })

  console.log(`Atualização finalizada.${tickets_fora_controles.length > 0 ? " Tickets fora dos controles: " + tickets_fora_controles : ""}`)
  
}

function atualizarDadosProtocolosAPI() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  var ultima_linha_dados_protocolos = planilha.getSheetByName("Dados dos protocolos").getRange("P2").getValue()
  
  var lin = 2
  var ticket_sem_info = planilha.getSheetByName("Tickets sem protocolo").getRange(lin,1).getValue()

  var tickets = []

  while (ticket_sem_info !== ""){
    tickets.push(ticket_sem_info)
    ticket_sem_info = planilha.getSheetByName("Tickets sem protocolo").getRange(++lin,1).getValue()
  }

  console.log("tickets sem protocolo: " + tickets)

  const fontes = [
    [//abc
      {
        "colunas": [0,1,2,11,12,12,35,36,37,38,39,0], // PROTOCOLO	DATA ENTRADA	DATA DE ABERTURA	CARGO	ESPECIALIDADE 1	ESPECIALIDADE 2	CARGA HORARIA	ESCALA	EQUIPE	UNIDADE	REGIONAL	FONTE
        "colunas_data": [1,2],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "GERAL",
        "fonte": "ABC"
      },
      "https://script.google.com/macros/s/AKfycbyCDiVFryD-0uz4VDaQ5QpbdNPTGlm8bAgVKdNKr935A5hvMJ-W1evShr8GWmDPMvfK/exec"
    ],
    [//cadm arq
      {
        "colunas": [0,1,4,5,6,7,8,9,10,11,12,0], // PROTOCOLO	MATRICULA	DATA INICIO
        "colunas_data": [1,4],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CONTRATAÇÕES ARQUIVADAS",
        "fonte": "CADM HISTORICO"
      },
      "https://script.google.com/macros/s/AKfycbzLpqp0v6b2mDGNTyD85XiBIBOUnEGFFcwIDNh8uzEec92sjHiT0vqpE1GUm6YK7FuZ/exec"
    ],
    [//cadm
      {
        "colunas": [0,1,4,5,6,7,8,9,10,11,12,0], // PROTOCOLO	MATRICULA	DATA INICIO
        "colunas_data": [1,4],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CDC",
        "fonte": "CDC"
      },
      "https://script.google.com/macros/s/AKfycbx1Y2QwYCpzkOpXaGqYSi2DLO7462r7mUfjTrTsxFlyhkfw61Glhxh7uqcfwEFXJQGcDw/exec"
    ],
  ]

  if(tickets.length == 0){
    console.log("Não há tickets para serem verificados")
    return
  }

  var csv = ""

  console.log('inciando as chamadas de API')
  
  fontes.forEach((busca) => {
      var options = {
        'method' : 'post', // Define o método como POST
        'contentType': 'application/json', // Informa que o corpo é JSON
        'payload' : JSON.stringify(busca[0]) // Converte o objeto para string JSON
      };

      var resposta = UrlFetchApp.fetch(busca[1],options)

      var texto_resposta = resposta.getContentText()

      if(texto_resposta !== ""){
        csv += texto_resposta + '\n'
      }

      console.log("finalizado " + busca[0].fonte)
    }
  )

  var dados = Utilities.parseCsv(csv);

  const tickets_dados = dados.map(dado => dado[0])

  var tickets_fora_controles = [] 
  tickets.forEach(ticket => {
    const consta = tickets_dados.indexOf(ticket) >= 0
    if(!consta){
      tickets_fora_controles.push(ticket)
    }
  })
  
  var aba = planilha.getSheetByName('Dados dos protocolos');

  console.log('iniciando inclusão dos dados das consultas')

  dados.forEach(dado => {
    const linha = ++ultima_linha_dados_protocolos
    aba.getRange(linha,1,1,dado.length).setValues([dado])
  })

  console.log(`Atualização finalizada.${tickets_fora_controles.length > 0 ? " Tickets fora dos controles: " + tickets_fora_controles : ""}`)
  
}

function atualizarFaseSituacao(){
  //pegar apenas os tickets que não estão encerrados
  // fazer uma comparação entre os tickets que estavam anteriormente como encerrados e verificar se estão com outra fase no relatório
}

function atualizarDadosContratacaoAPITeste() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  var ultima_linha_dados_contratacao = planilha.getSheetByName("Dados de contratação").getRange("G2").getValue()
  
  var lin = 2
  var ticket_sem_info = planilha.getSheetByName("Tickets sem info contratacao").getRange(lin,1).getValue()
  var ticket_sem_info_destino = planilha.getSheetByName("Tickets sem info contratacao").getRange(lin,3).getValue()

  var tickets = []
  var linhas_destino = []

  while (ticket_sem_info !== ""){
    tickets.push(ticket_sem_info)
    linhas_destino.push(ticket_sem_info_destino)
    
    lin++
    ticket_sem_info = planilha.getSheetByName("Tickets sem info contratacao").getRange(lin,1).getValue()
    ticket_sem_info_destino = planilha.getSheetByName("Tickets sem info contratacao").getRange(lin,3).getValue()
  }

  if(tickets.length == 0){
    console.log("Não há tickets para serem verificados")
    return
  }

  console.log("Busca iniciando com " + tickets.length + " tickets")

  const fontes = [
    [//abc
      {
        "colunas": [0, 9, 63, 0], // Converter arrays para string JSON
        "colunas_data": [63],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "GERAL",
        "fonte": "ABC"
      },
      "https://script.google.com/macros/s/AKfycbyCDiVFryD-0uz4VDaQ5QpbdNPTGlm8bAgVKdNKr935A5hvMJ-W1evShr8GWmDPMvfK/exec"
    ],
    [//ocupacao cadm
      {
        "colunas": [32, 9, 63, 0], // Converter arrays para string JSON
        "colunas_data": [63],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 33,
        "nome_aba_fonte": "GERAL",
        "fonte": "CADM OCUPADO POR ABC"
      },
      "https://script.google.com/macros/s/AKfycbyCDiVFryD-0uz4VDaQ5QpbdNPTGlm8bAgVKdNKr935A5hvMJ-W1evShr8GWmDPMvfK/exec"
    ],
    [//agendas
      {
        "colunas": [0, 1, 5, 0], // Converter arrays para string JSON
        "colunas_data": [5],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CONTRATADOS NA AGENDA",
        "fonte": "CONTRATADOS NAS AGENDAS"
      },
      "https://script.google.com/macros/s/AKfycbxoEJgUCrHbErS9kq-aBsF46etNKJyEZyhNZCWQOw2M3Gpjd0mHRVWc1BuOTcuEgh06/exec"
    ],
    [//cadm arq
      {
        "colunas": [0, 38, 43, 0], // PROTOCOLO	MATRICULA	DATA INICIO
        "colunas_data": [43],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CONTRATAÇÕES ARQUIVADAS",
        "fonte": "CADM HISTORICO"
      },
      "https://script.google.com/macros/s/AKfycbzLpqp0v6b2mDGNTyD85XiBIBOUnEGFFcwIDNh8uzEec92sjHiT0vqpE1GUm6YK7FuZ/exec"
    ],
    [//cadm
      {
        "colunas": [0, 38, 43, 0], // PROTOCOLO	MATRICULA	DATA INICIO
        "colunas_data": [43],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CDC",
        "fonte": "CDC"
      },
      "https://script.google.com/macros/s/AKfycbx1Y2QwYCpzkOpXaGqYSi2DLO7462r7mUfjTrTsxFlyhkfw61Glhxh7uqcfwEFXJQGcDw/exec"
    ],
  ]

  var csv = ""
  
  fontes.forEach((busca) => {
      var options = {
        'method' : 'post', // Define o método como POST
        'contentType': 'application/json', // Informa que o corpo é JSON
        'payload' : JSON.stringify(busca[0]) // Converte o objeto para string JSON
      };

      var resposta = UrlFetchApp.fetch(busca[1],options)

      var texto_resposta = resposta.getContentText()

      if(texto_resposta !== ""){
        csv += texto_resposta + '\n'
      }

      console.log("finalizado " + busca[0].fonte)
    }
  )

  var dados = Utilities.parseCsv(csv);

  const tickets_dados = dados.map(dado => dado[0])

  var tickets_fora_controles = [] 
  tickets.forEach(ticket => {
    if(ticket == '31.00462193/2025-76'){
      console.log("este está fora")
    }
    const consta = tickets_dados.indexOf(ticket) >= 0
    if(!consta){
      tickets_fora_controles.push(ticket)
    }
  })
  
  var aba = planilha.getSheetByName('Dados de contratação');

  console.log('iniciando inclusão dos dados das consultas')

  dados.forEach(dado => {
    
    const i = tickets.indexOf(dado[0])

    if(ticket_sem_info_destino < 0){
      console.log('aqui')
    }
    
    const linha = ticket_sem_info_destino < 0 ? ultima_linha_dados_contratacao++ : linhas_destino[i]
    aba.getRange(linha,1,1,dado.length).setValues([dado])
  })

  console.log(`Atualização finalizada.${tickets_fora_controles.length > 0 ? " Tickets fora dos controles: " + tickets_fora_controles : ""}`)
  
}

function atualizarDadosProtocolosAPITeste() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  var ultima_linha_dados_protocolos = planilha.getSheetByName("Dados dos protocolos").getRange("P2").getValue()
  
  var lin = 2
  var ticket_sem_info = planilha.getSheetByName("Tickets sem protocolo").getRange(lin,1).getValue()

  var tickets = []

  while (ticket_sem_info !== ""){
    tickets.push(ticket_sem_info)
    ticket_sem_info = planilha.getSheetByName("Tickets sem protocolo").getRange(++lin,1).getValue()
  }

  console.log("tickets sem protocolo: " + tickets)

  const fontes = [
    [//abc
      {
        "colunas": [0,1,2,11,12,12,35,36,37,38,39,0], // PROTOCOLO	DATA ENTRADA	DATA DE ABERTURA	CARGO	ESPECIALIDADE 1	ESPECIALIDADE 2	CARGA HORARIA	ESCALA	EQUIPE	UNIDADE	REGIONAL	FONTE
        "colunas_data": [1,2],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "GERAL",
        "fonte": "ABC"
      },
      "https://script.google.com/macros/s/AKfycbyCDiVFryD-0uz4VDaQ5QpbdNPTGlm8bAgVKdNKr935A5hvMJ-W1evShr8GWmDPMvfK/exec"
    ],
    [//cadm arq
      {
        "colunas": [0,1,4,5,6,7,8,9,10,11,12,0], // PROTOCOLO	MATRICULA	DATA INICIO
        "colunas_data": [1,4],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CONTRATAÇÕES ARQUIVADAS",
        "fonte": "CADM HISTORICO"
      },
      "https://script.google.com/macros/s/AKfycbzLpqp0v6b2mDGNTyD85XiBIBOUnEGFFcwIDNh8uzEec92sjHiT0vqpE1GUm6YK7FuZ/exec"
    ],
    [//cadm
      {
        "colunas": [0,1,4,5,6,7,8,9,10,11,12,0], // PROTOCOLO	MATRICULA	DATA INICIO
        "colunas_data": [1,4],
        "tickets": tickets, // Converter arrays para string JSON
        "coluna_ticket": 1,
        "nome_aba_fonte": "CDC",
        "fonte": "CDC"
      },
      "https://script.google.com/macros/s/AKfycbx1Y2QwYCpzkOpXaGqYSi2DLO7462r7mUfjTrTsxFlyhkfw61Glhxh7uqcfwEFXJQGcDw/exec"
    ],
  ]

  if(tickets.length == 0){
    console.log("Não há tickets para serem verificados")
    return
  }

  var csv = ""

  console.log('inciando as chamadas de API')
  
  fontes.forEach((busca) => {
      var options = {
        'method' : 'post', // Define o método como POST
        'contentType': 'application/json', // Informa que o corpo é JSON
        'payload' : JSON.stringify(busca[0]) // Converte o objeto para string JSON
      };

      var resposta = UrlFetchApp.fetch(busca[1],options)

      var texto_resposta = resposta.getContentText()

      if(texto_resposta !== ""){
        csv += texto_resposta + '\n'
      }

      console.log("finalizado " + busca[0].fonte)
    }
  )

  var dados = Utilities.parseCsv(csv);

  const tickets_dados = dados.map(dado => dado[0])

  var tickets_fora_controles = [] 
  tickets.forEach(ticket => {
    const consta = tickets_dados.indexOf(ticket) >= 0
    if(!consta){
      tickets_fora_controles.push(ticket)
    }
  })
  
  var aba = planilha.getSheetByName('Dados dos protocolos');

  console.log('iniciando inclusão dos dados das consultas')

  dados.forEach(dado => {
    const linha = ++ultima_linha_dados_protocolos
    aba.getRange(linha,1,1,dado.length).setValues([dado])
  })

  console.log(`Atualização finalizada.${tickets_fora_controles.length > 0 ? " Tickets fora dos controles: " + tickets_fora_controles : ""}`)
  
}

function atualizarFaseSituacaoTeste(){
  //pegar apenas os tickets que não estão encerrados
  // fazer uma comparação entre os tickets que estavam anteriormente como encerrados e verificar se estão com outra fase no relatório
}