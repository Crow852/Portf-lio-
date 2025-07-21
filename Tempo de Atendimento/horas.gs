function calcularHorasUteisAutomatico() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dados = aba.getDataRange().getValues();

  const idxDataCriacao = 1;
  const idxConclusao = 2;
  const idxResultado = 3;

  const ultimaLinha = aba.getLastRow();

  aba.getRange(2, idxDataCriacao + 1, ultimaLinha - 1).setNumberFormat("dd/MM/yyyy hh:mm:ss");
  aba.getRange(2, idxConclusao + 1, ultimaLinha - 1).setNumberFormat("dd/MM/yyyy hh:mm:ss");
  aba.getRange(2, idxResultado + 1, ultimaLinha - 1).setNumberFormat("[h]:mm");

  for (let i = 1; i < dados.length; i++) {
    const rawInicio = dados[i][idxDataCriacao];
    const rawFim = dados[i][idxConclusao];
    const dataInicio = new Date(rawInicio);
    const dataFim = rawFim ? new Date(rawFim) : new Date();

    if (!isNaN(dataInicio.getTime()) && !isNaN(dataFim.getTime())) {
      const minutosUteis = calcularMinutosUteis(dataInicio, dataFim);
      const horas = Math.floor(minutosUteis / 60);
      const minutos = minutosUteis % 60;
      const tempoFormatado = Utilities.formatString("%02d:%02d", horas, minutos);
      aba.getRange(i + 1, idxResultado + 1).setValue(tempoFormatado);
    }
  }
}

function calcularMinutosUteis(inicio, fim) {
  let totalMinutos = 0;
  let atual = new Date(inicio);
  atual.setSeconds(0);
  atual.setMilliseconds(0);

  while (atual <= fim) {
    const diaSemana = atual.getDay();
    if (diaSemana >= 1 && diaSemana <= 5) {
      const ehMesmoDia = atual.toDateString() === fim.toDateString();
      if (ehMesmoDia) {
        const diffMin = Math.floor((fim - atual) / (1000 * 60));
        totalMinutos += diffMin;
        break;
      } else {
        const minutosRestantesHoje = (24 * 60) - (atual.getHours() * 60 + atual.getMinutes());
        totalMinutos += minutosRestantesHoje;
      }
    }
    atual.setDate(atual.getDate() + 1);
    atual.setHours(0, 0, 0, 0);
  }
  return totalMinutos;
}

function enviarNotificacoesDiscord() {
  const webhook24h = "https://discordapp.com/api/webhooks/1371533574540361818/JcpnEE6Trt2UC3fmbYBXjbdhbz5xc_tyTzyEK_2ZTH1MsYZa2A41tICw2AZ51BY-4qtb";
  const webhook20h = "https://discordapp.com/api/webhooks/1372547885396131890/E8y1YT9V-Wuu3k7Rwh6CS49FjzlD1Jk1r1_DIvbalToFptSOzixGUFDmfh1Wc7Ba4A3C";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Licenciadas tempo atendimento");
  const backup = ss.getSheetByName("BackupNotificacoes");
  const dados = aba.getDataRange().getValues();
  const dadosBackup = backup.getDataRange().getValues();
  const agora = new Date();

  const idxID = dados[0].indexOf("ID Único");
  const idxNotificado = dadosBackup[0].indexOf("Notificado");
  const idxNotificado20h = dadosBackup[0].indexOf("Notificado 20h");
  const idxChave24 = dadosBackup[0].indexOf("Chave Notificação 24h");
  const idxChave20 = dadosBackup[0].indexOf("Chave Notificação 20h");

  const mapaBackup = new Map();
  for (let i = 1; i < dadosBackup.length; i++) {
    const id = dadosBackup[i][0];
    if (id) {
      mapaBackup.set(id.trim(), {
        linha: i + 1,
        notificado: dadosBackup[i][idxNotificado],
        notificado20h: dadosBackup[i][idxNotificado20h],
        chave24h: dadosBackup[i][idxChave24],
        chave20h: dadosBackup[i][idxChave20]
      });
    }
  }

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const licenciado = linha[0];
    const dataCriacao = new Date(linha[1]);
    if (isNaN(dataCriacao.getTime())) continue;

    const dataConclusao = linha[2] ? new Date(linha[2]) : null;
    const empresa = linha[4];
    const oportunidade = linha[5];
    const vendedor = linha[8];
    const status = (linha[10] || "").toString().toLowerCase().trim();
    const atendimentoConcluido = dataConclusao && !isNaN(dataConclusao.getTime());
    const statusFinalizado = status === "finalizada";

    const idUnico = linha[idxID];
    if (!idUnico || !mapaBackup.has(idUnico.trim())) continue;

    const fim = dataConclusao || agora;
    const minutosUteis = calcularMinutosUteis(dataCriacao, fim);
    const horasUteis = minutosUteis / 60;
    const chaveUnica = `${empresa}|${oportunidade}|${dataCriacao.toISOString().slice(0, 19)}`;
    const registroBackup = mapaBackup.get(idUnico.trim());
    const linhaBackup = registroBackup?.linha;

    const jaNotificado24 = registroBackup?.notificado === "Sim" || registroBackup?.chave24h === chaveUnica;
    const jaNotificado20 = registroBackup?.notificado20h === "Sim" || registroBackup?.chave20h === chaveUnica;

    if (horasUteis >= 24 && !jaNotificado24) {
      const msg24 = `\uD83D\uDCE2 Atendimento com atraso detectado, direcionar para outro vendedor \uD83D\uDCE2\n\n` +
        `Licenciado: ${licenciado}\n` +
        `Empresa/Conta: ${empresa}\n` +
        `Oportunidade: ${oportunidade}\n` +
        `Tempo limite excedido: + 24h úteis.`;

      UrlFetchApp.fetch(webhook24h, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ content: msg24 })
      });

      if (linhaBackup) {
        backup.getRange(linhaBackup, idxNotificado + 1).setValue("Sim");
        backup.getRange(linhaBackup, idxChave24 + 1).setValue(chaveUnica);
      }
      SpreadsheetApp.flush();
      Utilities.sleep(500);
    }

    if (horasUteis >= 20 && horasUteis < 24 && !atendimentoConcluido && !statusFinalizado && !jaNotificado20) {
      const msg20 = `\uD83D\uDD14 Atenção! Primeiro Contato Ainda Não Realizado\n\n` +
        `A oportunidade ${oportunidade}, vinculada à conta ${empresa}, segue sem atendimento após ${horasUteis.toFixed(2)} horas da sua criação.\n\n` +
        `Responsável: ${vendedor}\n\n` +
        `⚠️ O prazo máximo para o primeiro contato é de 24 horas. Restam apenas ${(24 - horasUteis).toFixed(2)} horas para evitar o descumprimento do SLA.`;

      UrlFetchApp.fetch(webhook20h, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ content: msg20 })
      });

      if (linhaBackup) {
        backup.getRange(linhaBackup, idxNotificado20h + 1).setValue("Sim");
        backup.getRange(linhaBackup, idxChave20 + 1).setValue(chaveUnica);
      }
      SpreadsheetApp.flush();
      Utilities.sleep(500);
    }
  }
}

function sincronizarNotificacoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Licenciadas tempo atendimento");
  const backup = ss.getSheetByName("BackupNotificacoes");
  const dados = aba.getDataRange().getValues();
  const dadosBackup = backup.getDataRange().getValues();

  const idxID = dados[0].indexOf("ID Único");
  const idxNotificacao = dados[0].indexOf("Notificado");

  const mapaBackup = new Map();
  for (let i = 1; i < dadosBackup.length; i++) {
    const [id, ...valores] = dadosBackup[i];
    mapaBackup.set(id.trim(), valores);
  }

  for (let i = 1; i < dados.length; i++) {
    const id = dados[i][idxID];
    if (id && mapaBackup.has(id.trim())) {
      const valores = mapaBackup.get(id.trim());
      for (let j = 0; j < valores.length; j++) {
        aba.getRange(i + 1, idxNotificacao + 1 + j).setValue(valores[j]);
      }
    }
  }
}

function atualizarIDUnico() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Licenciadas tempo atendimento");
  const dados = sheet.getDataRange().getValues();
  const colunaIDUnico = 16;

  for (let i = 1; i < dados.length; i++) {
    const empresa = dados[i][4];
    const oportunidade = dados[i][5];
    const dataCriacaoRaw = dados[i][1];
    if (empresa && oportunidade && dataCriacaoRaw) {
      const dataCriacao = new Date(dataCriacaoRaw);
      const idUnico = `${empresa}|${oportunidade}|${dataCriacao.toISOString().slice(0, 19)}`;
      sheet.getRange(i + 1, colunaIDUnico + 1).setValue(idUnico);
    }
  }
}

function salvarBackupNotificacoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Licenciadas tempo atendimento");
  const backup = ss.getSheetByName("BackupNotificacoes") || ss.insertSheet("BackupNotificacoes");

  const dados = aba.getDataRange().getValues();
  const header = dados[0];
  const idIndex = header.indexOf("ID Único");
  const notifIndex = header.indexOf("Notificado");

  const headerBackup = ["ID Único", ...header.slice(notifIndex)];
  const dadosBackup = [headerBackup];

  for (let i = 1; i < dados.length; i++) {
    const id = dados[i][idIndex];
    if (id) {
      const notificacoes = dados[i].slice(notifIndex);
      dadosBackup.push([id, ...notificacoes]);
    }
  }

  backup.clearContents();
  backup.getRange(1, 1, dadosBackup.length, dadosBackup[0].length).setValues(dadosBackup);
}

function atualizarPlanilhaCompletamente() {
  salvarBackupNotificacoes();
  calcularHorasUteisAutomatico();
  atualizarIDUnico();
  enviarNotificacoesDiscord();
  sincronizarNotificacoes();
}
