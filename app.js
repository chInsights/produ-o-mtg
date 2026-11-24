const firebaseConfig = {
  apiKey: "AIzaSyD8xIc7gLsG_u2E4GjZ1a9lA2_zUsvbs04",
  authDomain: "pcp-mtg-fab.firebaseapp.com",
  projectId: "pcp-mtg-fab",
  storageBucket: "pcp-mtg-fab.firebasestorage.app",
  messagingSenderId: "749171419070",
  appId: "1:749171419070:web:a0c92919c443f5c1f1fbef"
};

const COLLECTION_NAME = "pcp_mtg_fab";
let db = null;

let rawPlanoRows = [];
let sapMap = new Map();
let observacoesMap = {}; // Mapa leve: { "WO_1000018613": "Falta vidro", ... }
let consolidatedOrders = [];
let currentTab = 'todas';

// =========================================================================
// 1. CONTROLE DE SELETOR DE ARQUIVOS E DRAG & DROP
// =========================================================================
function abrirSeletor(inputId) {
  const inp = document.getElementById(inputId);
  if (inp) inp.click();
}

function setupDragAndDrop(cardId, handlerFn) {
  const card = document.getElementById(cardId);
  if (!card) return;
  card.addEventListener('dragover', (e) => {
    e.preventDefault();
    card.classList.add('dragover');
  });
  card.addEventListener('dragleave', () => card.classList.remove('dragover'));
  card.addEventListener('drop', (e) => {
    e.preventDefault();
    card.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handlerFn(e.dataTransfer.files[0]);
    }
  });
}

setupDragAndDrop('cardPlano', processarArquivoPlano);
setupDragAndDrop('cardSap', processarArquivoSap);

document.getElementById('inputPlano').addEventListener('change', (e) => {
  if (e.target.files && e.target.files[0]) {
    processarArquivoPlano(e.target.files[0]);
    e.target.value = '';
  }
});

document.getElementById('inputSap').addEventListener('change', (e) => {
  if (e.target.files && e.target.files[0]) {
    processarArquivoSap(e.target.files[0]);
    e.target.value = '';
  }
});

// =========================================================================
// 2. CONTROLES DE INTERFACE
// =========================================================================
const refDateInput = document.getElementById('refDateInput');
refDateInput.value = new Date().toISOString().split('T')[0];
refDateInput.addEventListener('change', processAndRender);

const searchInput = document.getElementById('searchInput');
const btnClearSearch = document.getElementById('btnClearSearch');

searchInput.addEventListener('input', () => {
  btnClearSearch.style.display = searchInput.value.trim() ? 'inline-block' : 'none';
  renderTable();
});

function limparBusca() {
  searchInput.value = '';
  btnClearSearch.style.display = 'none';
  searchInput.focus();
  renderTable();
}

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentTab = btn.dataset.tab;
    document.getElementById('btnCopiarEncerrar').style.display = currentTab === 'encerrar' ? 'inline-flex' : 'none';
    renderTable();
  });
});

// =========================================================================
// 3. FUNÇÕES AUXILIARES DE DATA, NÚMERO E TEXTO
// =========================================================================
function parseExcelDate(val) {
  if (!val) return null;
  if (typeof val === 'number' && val > 20000) {
    const utcDays = Math.floor(val - 25569);
    const dateObj = new Date(utcDays * 86400 * 1000);
    return dateObj.toISOString().split('T')[0];
  }
  const str = String(val).trim();
  const brMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (brMatch) {
    return `${brMatch[3]}-${brMatch[2].padStart(2, '0')}-${brMatch[1].padStart(2, '0')}`;
  }
  const brShort = str.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (brShort) {
    const year = new Date().getFullYear();
    return `${year}-${brShort[2].padStart(2, '0')}-${brShort[1].padStart(2, '0')}`;
  }
  const meses = { jan:'01', fev:'02', mar:'03', abr:'04', mai:'05', jun:'06', jul:'07', ago:'08', set:'09', out:'10', nov:'11', dez:'12' };
  const ptMatch = str.toLowerCase().match(/^(\d{1,2})[-\/.]([a-z]{3})$/);
  if (ptMatch && meses[ptMatch[2]]) {
    const year = new Date().getFullYear();
    return `${year}-${meses[ptMatch[2]]}-${ptMatch[1].padStart(2, '0')}`;
  }
  return null;
}

function formatDateBR(iso) {
  if (!iso || iso === 'SEM_DATA') return 'S/ Data';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}`;
}

function parseNumber(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  const str = String(val).trim();
  if (str === '-' || str === '') return 0;
  if (str.includes(',') && str.includes('.')) {
    return parseFloat(str.replace(/\./g, '').replace(',', '.')) || 0;
  }
  if (str.includes(',')) {
    return parseFloat(str.replace(',', '.')) || 0;
  }
  if (/^\d{1,3}(\.\d{3})+$/.test(str)) {
    return parseFloat(str.replace(/\./g, '')) || 0;
  }
  return parseFloat(str.replace(/[^\d.-]/g, '')) || 0;
}

function normalizeKey(str) {
  return String(str || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function getNowBR() {
  return new Date().toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
}

// =========================================================================
// 4. LEITURA DAS PLANILHAS (PLANO E SAP)
// =========================================================================
function processarArquivoPlano(file) {
  if (typeof XLSX === 'undefined') {
    alert('A biblioteca de leitura Excel (SheetJS) não carregou. Recarregue a página (Ctrl+F5).');
    return;
  }

  document.getElementById('planoStatus').textContent = 'Lendo...';
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const wb = XLSX.read(evt.target.result, { type: 'array' });
      const sheetName = wb.SheetNames.find(n => n.toUpperCase().includes('MONTAGEM')) || wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      let headerIdx = matrix.findIndex(row =>
        row.some(cell => normalizeKey(cell) === 'wo') &&
        row.some(cell => normalizeKey(cell) === 'item')
      );

      if (headerIdx === -1) {
        alert('Não encontrei as colunas "Wo" e "Item" na aba ' + sheetName);
        document.getElementById('planoStatus').textContent = 'Erro';
        return;
      }

      const headers = matrix[headerIdx].map(h => normalizeKey(h));
      const col = {
        linha: headers.findIndex(h => h === 'linha'),
        data: headers.findIndex(h => h === 'data'),
        wo: headers.findIndex(h => h === 'wo'),
        item: headers.findIndex(h => h === 'item'),
        obs: headers.findIndex(h => h.includes('observa')),
        qtd: headers.findIndex(h => h === 'qtd'),
        classe: headers.findIndex(h => h.includes('class')),
        cliente: headers.findIndex(h => h.includes('cliente')),
        horas: headers.findIndex(h => h === 'horas' || h.includes('horas')),
        demandas: headers.findIndex(h => h.includes('demanda')),
        duzias: headers.findIndex(h => h.includes('duzia'))
      };

      const parsedRows = [];
      let blocoAtual = 'OFICIAL';

      for (let i = headerIdx + 1; i < matrix.length; i++) {
        const r = matrix[i];
        const colA = normalizeKey(r[col.linha >= 0 ? col.linha : 0]);

        if (colA.includes('considerar')) {
          blocoAtual = 'CONSIDERAR';
          continue;
        }

        let wo = String(r[col.wo] || '').trim();
        if (wo === '-' || wo === '0') wo = '';

        const item = String(r[col.item] || '').trim();
        const qtd = parseNumber(r[col.qtd]);
        const horas = col.horas >= 0 ? parseNumber(r[col.horas]) : 0;
        const duzias = col.duzias >= 0 ? parseNumber(r[col.duzias]) : 0;
        const dataIso = parseExcelDate(r[col.data]) || 'SEM_DATA';

        if (!item || qtd <= 0) continue;

        parsedRows.push({
          rowIndex: i + 1,
          bloco: blocoAtual,
          wo,
          item,
          dataIso,
          obs: String(r[col.obs] || '').trim(),
          qtd,
          horas,
          duzias,
          classe: String(r[col.classe] || '').trim(),
          cliente: String(r[col.cliente] || '').trim(),
          demanda: col.demandas >= 0 ? String(r[col.demandas] || '').trim() : ''
        });
      }

      rawPlanoRows = parsedRows;
      // Limpa localmente as observações ao carregar nova planilha mensal
      observacoesMap = {};

      const totalCxLido = rawPlanoRows.reduce((acc, r) => acc + r.qtd, 0);

      document.getElementById('cardPlano').classList.add('loaded');
      document.getElementById('planoFileName').textContent = `${file.name} (${rawPlanoRows.length} linhas • ${totalCxLido.toLocaleString('pt-BR')} cx)`;
      document.getElementById('planoStatus').textContent = 'Carregado ✓';

      processAndRender();
      savePlanoToCloud(file.name, parsedRows);
    } catch (err) {
      console.error(err);
      alert('Erro ao ler a planilha do Plano: ' + err.message);
      document.getElementById('planoStatus').textContent = 'Erro';
    }
  };
  reader.readAsArrayBuffer(file);
}

function processarArquivoSap(file) {
  if (typeof XLSX === 'undefined') {
    alert('A biblioteca de leitura Excel (SheetJS) não carregou. Recarregue a página (Ctrl+F5).');
    return;
  }

  document.getElementById('sapStatus').textContent = 'Lendo...';
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const wb = XLSX.read(evt.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      let headerIdx = matrix.findIndex(row =>
        row.some(cell => normalizeKey(cell) === 'ordem') &&
        row.some(cell => normalizeKey(cell).includes('boa confirmada'))
      );

      if (headerIdx === -1) {
        alert('Não encontrei as colunas "Ordem" e "Quantidade boa confirmada" no relatório SAP.');
        document.getElementById('sapStatus').textContent = 'Erro';
        return;
      }

      const headers = matrix[headerIdx].map(h => normalizeKey(h));
      const col = {
        ordem: headers.findIndex(h => h === 'ordem'),
        material: headers.findIndex(h => h === 'material'),
        desc: headers.findIndex(h => h.includes('texto breve')),
        qtdOrdem: headers.findIndex(h => h.includes('quantidade da ordem')),
        qtdBoa: headers.findIndex(h => h.includes('boa confirmada'))
      };

      const newMap = new Map();
      const plainObj = {};

      for (let i = headerIdx + 1; i < matrix.length; i++) {
        const r = matrix[i];
        const ordem = String(r[col.ordem] || '').trim();
        if (!ordem) continue;

        const entry = {
          material: String(r[col.material] || '').trim(),
          desc: String(r[col.desc] || '').trim(),
          qtdOrdemSap: parseNumber(r[col.qtdOrdem]),
          qtdBoa: parseNumber(r[col.qtdBoa])
        };
        newMap.set(ordem, entry);
        plainObj[ordem] = entry;
      }

      sapMap = newMap;

      document.getElementById('cardSap').classList.add('loaded');
      document.getElementById('sapFileName').textContent = `${file.name} (${sapMap.size} ordens SAP)`;
      document.getElementById('sapStatus').textContent = 'Atualizado ✓';

      processAndRender();
      saveSapToCloud(file.name, plainObj);
    } catch (err) {
      console.error(err);
      alert('Erro ao ler o relatório SAP: ' + err.message);
      document.getElementById('sapStatus').textContent = 'Erro';
    }
  };
  reader.readAsArrayBuffer(file);
}

// =========================================================================
// 5. FIREBASE FIRESTORE (PLANO, SAP E OBSERVAÇÕES LEVES)
// =========================================================================
async function savePlanoToCloud(fileName, rows) {
  if (!db) return;
  const badge = document.getElementById('cloudBadge');
  const meta = document.getElementById('planoCloudMeta');
  badge.className = "cloud-badge syncing";
  badge.textContent = "🔄 Salvando plano...";
  meta.textContent = "Enviando para a nuvem...";

  try {
    const CHUNK_SIZE = 1000;
    const totalChunks = Math.ceil(rows.length / CHUNK_SIZE) || 1;
    const nowStr = getNowBR();

    for (let c = 0; c < totalChunks; c++) {
      const slice = rows.slice(c * CHUNK_SIZE, (c + 1) * CHUNK_SIZE);
      await db.collection(COLLECTION_NAME).doc(`plano_chunk_${c}`).set({ rows: slice });
    }

    await db.collection(COLLECTION_NAME).doc("plano_meta").set({
      fileName,
      updatedAt: nowStr,
      totalRows: rows.length,
      totalChunks
    });

    // Reset automático das observações do mês anterior ao subir novo plano
    await db.collection(COLLECTION_NAME).doc("observacoes").set({
      notas: {},
      resetAt: nowStr
    });

    badge.className = "cloud-badge online";
    badge.textContent = "☁️ Nuvem Sincronizada";
    meta.textContent = `☁️ Salvo na nuvem em ${nowStr} (Obs. resetadas)`;
  } catch (err) {
    console.error("Erro ao salvar plano:", err);
    badge.className = "cloud-badge offline";
    badge.textContent = "⚠️ Falha ao salvar na nuvem";
    meta.textContent = `Erro nuvem: ${err.code || err.message}`;
  }
}

async function saveSapToCloud(fileName, entriesObj) {
  if (!db) return;
  const badge = document.getElementById('cloudBadge');
  const meta = document.getElementById('sapCloudMeta');
  badge.className = "cloud-badge syncing";
  badge.textContent = "🔄 Salvando SAP...";
  meta.textContent = "Enviando para a nuvem...";

  try {
    const nowStr = getNowBR();
    await db.collection(COLLECTION_NAME).doc("sap_atual").set({
      fileName,
      updatedAt: nowStr,
      entries: entriesObj
    });

    badge.className = "cloud-badge online";
    badge.textContent = "☁️ Nuvem Sincronizada";
    meta.textContent = `☁️ Atualizado na nuvem em ${nowStr}`;
  } catch (err) {
    console.error("Erro ao salvar SAP:", err);
    badge.className = "cloud-badge offline";
    badge.textContent = "⚠️ Falha ao salvar na nuvem";
    meta.textContent = `Erro nuvem: ${err.code || err.message}`;
  }
}

async function salvarObservacao(groupKey, inputEl) {
  const texto = String(inputEl.value || '').trim();
  const anterior = observacoesMap[groupKey] || '';

  // Visual instantâneo no campo
  if (texto) {
    inputEl.classList.add('has-note');
  } else {
    inputEl.classList.remove('has-note');
  }

  // Se não houve alteração real, não gasta escrita no banco
  if (texto === anterior) return;

  if (texto) {
    observacoesMap[groupKey] = texto;
  } else {
    delete observacoesMap[groupKey];
  }

  // Atualiza também no array em memória para filtro/exportação
  const ordObj = consolidatedOrders.find(o => o.groupKey === groupKey);
  if (ordObj) ordObj.notaPcp = texto;

  if (!db) return;

  try {
    await db.collection(COLLECTION_NAME).doc("observacoes").set({
      notas: observacoesMap,
      updatedAt: getNowBR()
    });

    const icon = inputEl.parentElement.querySelector('.obs-saved-icon');
    if (icon) {
      icon.classList.add('show');
      setTimeout(() => icon.classList.remove('show'), 1400);
    }
  } catch (err) {
    console.error("Erro ao salvar observação:", err);
  }
}

async function initCloudSync() {
  const badge = document.getElementById('cloudBadge');
  if (typeof firebase === 'undefined') {
    badge.className = "cloud-badge offline";
    badge.textContent = "⚠️ Modo Local (SDK bloqueado)";
    return;
  }

  try {
    firebase.initializeApp(firebaseConfig);
    db = firebase.firestore();

    const metaSnap = await db.collection(COLLECTION_NAME).doc("plano_meta").get();
    if (metaSnap.exists) {
      const meta = metaSnap.data();
      let allRows = [];
      for (let c = 0; c < (meta.totalChunks || 1); c++) {
        const chunkSnap = await db.collection(COLLECTION_NAME).doc(`plano_chunk_${c}`).get();
        if (chunkSnap.exists) {
          allRows = allRows.concat(chunkSnap.data().rows || []);
        }
      }
      if (allRows.length > 0 && rawPlanoRows.length === 0) {
        rawPlanoRows = allRows;
        const totalCx = rawPlanoRows.reduce((acc, r) => acc + r.qtd, 0);
        document.getElementById('cardPlano').classList.add('loaded');
        document.getElementById('planoFileName').textContent = `${meta.fileName} (${rawPlanoRows.length} linhas • ${totalCx.toLocaleString('pt-BR')} cx)`;
        document.getElementById('planoCloudMeta').textContent = `☁️ Carregado da nuvem (${meta.updatedAt})`;
        document.getElementById('planoStatus').textContent = 'Na Nuvem ✓';
      }
    } else {
      document.getElementById('planoCloudMeta').textContent = "Nenhum plano salvo ainda. Selecione o Excel acima.";
    }

    const sapSnap = await db.collection(COLLECTION_NAME).doc("sap_atual").get();
    if (sapSnap.exists && sapMap.size === 0) {
      const data = sapSnap.data();
      sapMap = new Map(Object.entries(data.entries || {}));
      document.getElementById('cardSap').classList.add('loaded');
      const nomeSap = data.fileName || 'COOIS_1502_ATUAL.xlsx';
      document.getElementById('sapFileName').textContent = `${nomeSap} (${sapMap.size} ordens SAP)`;
      document.getElementById('sapCloudMeta').textContent = `☁️ Carregado da nuvem (${data.updatedAt})`;
      document.getElementById('sapStatus').textContent = 'Na Nuvem ✓';
    } else if (!sapSnap.exists) {
      document.getElementById('sapCloudMeta').textContent = "Nenhuma extração SAP salva ainda.";
    }

    // Carrega o documento leve de observações
    const obsSnap = await db.collection(COLLECTION_NAME).doc("observacoes").get();
    if (obsSnap.exists) {
      observacoesMap = obsSnap.data().notas || {};
    }

    badge.className = "cloud-badge online";
    badge.textContent = "☁️ Nuvem Conectada";
    processAndRender();
  } catch (err) {
    console.error("Erro ao conectar Firebase:", err);
    badge.className = "cloud-badge offline";
    badge.textContent = `⚠️ Modo Local (${err.code || 'sem nuvem'})`;
  }
}

// =========================================================================
// 6. MOTOR DE CRUZAMENTO USANDO "QUANTIDADE DA ORDEM (GMEIN)"
// =========================================================================
function processAndRender() {
  if (rawPlanoRows.length === 0) return;

  const refDate = refDateInput.value;
  const blocoFilter = document.getElementById('filterBloco').value;
  const classFilter = document.getElementById('filterClass').value;

  const filteredRows = rawPlanoRows.filter(r => {
    if (blocoFilter !== 'TODOS' && r.bloco !== blocoFilter) return false;
    if (classFilter !== 'TODOS' && r.classe.toUpperCase() !== classFilter) return false;
    return true;
  });

  const grouped = new Map();

  filteredRows.forEach((row) => {
    const groupKey = row.wo ? `WO_${row.wo}` : `SEM_WO_${row.item}_${row.dataIso}_${row.rowIndex}`;

    if (!grouped.has(groupKey)) {
      grouped.set(groupKey, {
        groupKey,
        wo: row.wo || '',
        hasWo: Boolean(row.wo),
        bloco: row.bloco,
        item: row.item,
        classe: row.classe,
        obs: row.obs,
        cliente: row.cliente,
        demanda: row.demanda,
        datas: [],
        progTotal: 0,
        progAteData: 0,
        duziasTotal: 0,
        duziasAteData: 0,
        horasTotal: 0,
        horasAteData: 0
      });
    }

    const entry = grouped.get(groupKey);
    entry.progTotal += row.qtd;
    entry.duziasTotal += row.duzias;
    entry.horasTotal += row.horas;

    if (row.dataIso !== 'SEM_DATA' && row.dataIso <= refDate) {
      entry.progAteData += row.qtd;
      entry.duziasAteData += row.duzias;
      entry.horasAteData += row.horas;
    }
    entry.datas.push({ dataIso: row.dataIso, qtd: row.qtd });
  });

  consolidatedOrders = [];
  let totalProgAteHoje = 0;
  let totalProgPlano = 0;
  let totalDuziasAteHoje = 0;
  let totalDuziasPlano = 0;
  let totalHorasAteHoje = 0;
  let totalHorasPlano = 0;
  let totalRealizado = 0;
  let deficitAtraso = 0;

  grouped.forEach(ord => {
    ord.datas.sort((a, b) => a.dataIso.localeCompare(b.dataIso));
    const sapInfo = ord.hasWo ? sapMap.get(ord.wo) : null;
    const realizado = sapInfo ? sapInfo.qtdBoa : 0;

    const qtdOrdemAlvo = (sapInfo && sapInfo.qtdOrdemSap > 0) ? sapInfo.qtdOrdemSap : ord.progTotal;
    const hasSapOrder = Boolean(sapInfo && sapInfo.qtdOrdemSap > 0);

    const desc = (sapInfo && sapInfo.desc) ? sapInfo.desc : (ord.obs || ord.demanda || 'Sem descrição SAP');
    const saldo = Math.max(0, qtdOrdemAlvo - realizado);
    const pct = qtdOrdemAlvo > 0 ? Math.min(100, Math.round((realizado / qtdOrdemAlvo) * 100)) : 0;

    const jaRealizadoAntesDoPlano = Math.max(0, qtdOrdemAlvo - ord.progTotal);
    const metaAcumuladaAteData = jaRealizadoAntesDoPlano + ord.progAteData;

    let status = 'prazo';
    const datasValidas = ord.datas.filter(d => d.dataIso !== 'SEM_DATA');
    const primeiraData = datasValidas.length > 0 ? datasValidas[0].dataIso : '9999-12-31';
    const ultimaData = datasValidas.length > 0 ? datasValidas[datasValidas.length - 1].dataIso : '9999-12-31';

    if (!ord.hasWo) {
      status = 'sem_wo';
    } else if (realizado >= qtdOrdemAlvo && qtdOrdemAlvo > 0) {
      status = 'encerrar';
    } else if (realizado < metaAcumuladaAteData && ultimaData < refDate) {
      status = 'atrasadas';
      deficitAtraso += (metaAcumuladaAteData - realizado);
    } else if (realizado < metaAcumuladaAteData && primeiraData <= refDate) {
      status = primeiraData < refDate ? 'atrasadas' : 'atencao';
      if (status === 'atrasadas') deficitAtraso += (metaAcumuladaAteData - realizado);
    } else if (datasValidas.some(d => d.dataIso === refDate)) {
      status = 'atencao';
    }

    totalProgAteHoje += ord.progAteData;
    totalProgPlano += ord.progTotal;
    totalDuziasAteHoje += ord.duziasAteData;
    totalDuziasPlano += ord.duziasTotal;
    totalHorasAteHoje += ord.horasAteData;
    totalHorasPlano += ord.horasTotal;
    totalRealizado += realizado;

    consolidatedOrders.push({
      ...ord,
      qtdOrdemAlvo,
      hasSapOrder,
      desc,
      realizado,
      saldo,
      pct,
      status,
      notaPcp: observacoesMap[ord.groupKey] || ''
    });
  });

  const pesoStatus = { atrasadas: 1, atencao: 2, prazo: 3, sem_wo: 4, encerrar: 5 };
  consolidatedOrders.sort((a, b) => pesoStatus[a.status] - pesoStatus[b.status] || a.datas[0].dataIso.localeCompare(b.datas[0].dataIso));

  document.getElementById('kpiProgTotal').textContent = `${totalProgPlano.toLocaleString('pt-BR')} cx`;
  document.getElementById('kpiProgAteHoje').textContent = `Até Data Base: ${totalProgAteHoje.toLocaleString('pt-BR')} cx`;

  document.getElementById('kpiDuziasTotal').textContent = `${Math.round(totalDuziasPlano).toLocaleString('pt-BR')} dz`;
  document.getElementById('kpiDuziasAteHoje').textContent = `Até Data Base: ${Math.round(totalDuziasAteHoje).toLocaleString('pt-BR')} dz`;

  document.getElementById('kpiHorasTotal').textContent = `${totalHorasPlano.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} h`;
  document.getElementById('kpiHorasAteHoje').textContent = `Até Data Base: ${totalHorasAteHoje.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} h`;

  document.getElementById('kpiRealizado').textContent = `${totalRealizado.toLocaleString('pt-BR')} cx`;
  const aderencia = totalProgAteHoje > 0 ? Math.round((totalRealizado / totalProgAteHoje) * 100) : 0;
  document.getElementById('kpiAderencia').textContent = `Relação vs Prog. Data: ${aderencia}%`;

  const qtdAtrasadas = consolidatedOrders.filter(o => o.status === 'atrasadas').length;
  const qtdEncerrar = consolidatedOrders.filter(o => o.status === 'encerrar').length;
  document.getElementById('kpiAtrasadas').textContent = qtdAtrasadas;
  document.getElementById('kpiSaldoAtraso').textContent = `Déficit acumulado: ${deficitAtraso.toLocaleString('pt-BR')} cx`;
  document.getElementById('kpiEncerrar').textContent = qtdEncerrar;

  document.getElementById('cntTodas').textContent = consolidatedOrders.length;
  document.getElementById('cntAtrasadas').textContent = qtdAtrasadas;
  document.getElementById('cntAtencao').textContent = consolidatedOrders.filter(o => o.status === 'atencao').length;
  document.getElementById('cntPrazo').textContent = consolidatedOrders.filter(o => o.status === 'prazo').length;
  document.getElementById('cntSemWo').textContent = consolidatedOrders.filter(o => o.status === 'sem_wo').length;
  document.getElementById('cntEncerrar').textContent = qtdEncerrar;

  renderTable();
}

function renderTable() {
  const tbody = document.getElementById('tableBody');
  const rawQuery = normalizeKey(searchInput.value);
  const terms = rawQuery.split(/[,;]+/).map(t => t.trim()).filter(Boolean);

  const filtered = consolidatedOrders.filter(o => {
    if (currentTab !== 'todas' && o.status !== currentTab) return false;
    if (terms.length > 0) {
      const hay = normalizeKey(`${o.wo} ${o.item} ${o.desc} ${o.cliente} ${o.demanda} ${o.notaPcp}`);
      const matched = terms.some(term => hay.includes(term));
      if (!matched) return false;
    }
    return true;
  });

  const counterEl = document.getElementById('searchCounter');
  if (consolidatedOrders.length > 0) {
    const somaCxFiltro = filtered.reduce((acc, o) => acc + o.progTotal, 0);
    const somaDzFiltro = filtered.reduce((acc, o) => acc + o.duziasTotal, 0);
    counterEl.textContent = `${filtered.length} itens (${somaCxFiltro.toLocaleString('pt-BR')} cx • ${Math.round(somaDzFiltro).toLocaleString('pt-BR')} dz)`;
  } else {
    counterEl.textContent = '';
  }

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="13" class="empty-state">Nenhuma ordem ou artigo encontrado neste filtro.</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map(o => {
    const badgeMap = {
      atrasadas: `<span class="pill pill-danger">🔴 Atrasada</span>`,
      atencao: `<span class="pill pill-warning">🟡 Atenção / Hoje</span>`,
      prazo: `<span class="pill pill-info">🔵 No Prazo</span>`,
      sem_wo: `<span class="pill pill-mts">⚪ Sem Wo</span>`,
      encerrar: `<span class="pill pill-success">🟢 Encerrar (100%)</span>`
    };
    const barColor = o.status === 'encerrar' ? 'var(--success)' : o.status === 'atrasadas' ? 'var(--danger)' : 'var(--accent)';
    const datasStr = o.datas.map(d => `${formatDateBR(d.dataIso)} (${d.qtd.toLocaleString('pt-BR')})`).join(' • ');
    const classBadge = o.classe.includes('MTO') ? `<span class="pill pill-mto">MTO</span>` : `<span class="pill pill-mts">MTS</span>`;
    const blocoBadge = o.bloco === 'CONSIDERAR' ? `<span class="pill pill-purple" style="margin-top:3px;">Considerar</span>` : '';

    const diffOrdemPlano = o.hasSapOrder && o.qtdOrdemAlvo !== o.progTotal;
    const sapOrdemCol = o.hasSapOrder
      ? `<span class="mono" style="${diffOrdemPlano ? 'color:var(--warning); font-weight:800;' : ''}">${o.qtdOrdemAlvo.toLocaleString('pt-BR')}</span>`
      : `<span style="color:var(--text-muted); font-size:0.76rem;">—</span>`;

    const notaAtual = escapeHtml(o.notaPcp);
    const notaClass = o.notaPcp ? 'obs-input has-note' : 'obs-input';

    return `
      <tr>
        <td>${badgeMap[o.status]}</td>
        <td class="mono">${o.hasWo ? o.wo : '<span style="color:var(--text-muted); font-weight:500;">—</span>'}</td>
        <td>
          <div style="font-weight:700; color:var(--text-main);">${o.item}</div>
          <div style="font-size:0.75rem; color:var(--text-muted);">${o.desc}</div>
          ${o.demanda ? `<div style="font-size:0.71rem; color:var(--accent); font-weight:600; margin-top:2px;">${o.demanda}</div>` : ''}
        </td>
        <td>
          <div>${classBadge}</div>
          ${blocoBadge}
        </td>
        <td style="font-size:0.77rem; color:var(--text-muted); max-width:200px;">${datasStr}</td>
        <td class="mono">${o.progTotal.toLocaleString('pt-BR')}</td>
        <td>${sapOrdemCol}</td>
        <td class="mono" style="color:var(--purple);">${Math.round(o.duziasTotal).toLocaleString('pt-BR')}</td>
        <td class="mono">${o.horasTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} h</td>
        <td class="mono" style="color:var(--info);">${o.realizado.toLocaleString('pt-BR')}</td>
        <td class="mono" style="color:${o.saldo > 0 ? 'var(--danger)' : 'var(--success)'};">${o.saldo.toLocaleString('pt-BR')}</td>
        <td>
          <div class="mono" style="font-size:0.77rem;">${o.pct}%</div>
          <div class="progress-bar"><div class="progress-fill" style="width:${o.pct}%; background:${barColor};"></div></div>
        </td>
        <td class="obs-cell">
          <div class="obs-input-wrap">
            <input
              type="text"
              class="${notaClass}"
              value="${notaAtual}"
              placeholder="+ Obs..."
              title="${notaAtual || 'Clique para adicionar observação'}"
              onblur="salvarObservacao('${o.groupKey}', this)"
              onkeydown="if(event.key==='Enter'){this.blur();}"
            />
            <span class="obs-saved-icon">✓</span>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function copiarOrdensEncerrar() {
  const lista = consolidatedOrders.filter(o => o.status === 'encerrar' && o.hasWo).map(o => o.wo).join('\n');
  if (!lista) return alert('Não há ordens 100% concluídas para copiar.');
  navigator.clipboard.writeText(lista);
  alert('Lista de ordens (Wo) copiada! Cole direto na seleção múltipla da COOIS no SAP.');
}

function exportarExcel(tipo) {
  if (consolidatedOrders.length === 0) {
    return alert('Importe o Plano de Produção e o Relatório SAP antes de exportar.');
  }

  const lista = tipo === 'finalizadas'
    ? consolidatedOrders.filter(o => o.status === 'encerrar')
    : consolidatedOrders.filter(o => o.status !== 'encerrar');

  if (lista.length === 0) {
    return alert(tipo === 'finalizadas'
      ? 'Não há ordens finalizadas (100%) para exportar.'
      : 'Não há ordens pendentes para exportar.');
  }

  const labelStatus = {
    atrasadas: 'ATRASADA',
    atencao: 'ATENÇÃO / HOJE',
    prazo: 'NO PRAZO',
    sem_wo: 'SEM WO / PREVISÃO',
    encerrar: 'CONCLUÍDA (ENCERRAR SAP)'
  };

  const rows = lista.map(o => ({
    'Status': labelStatus[o.status],
    'Bloco': o.bloco,
    'Ordem (Wo)': o.wo || 'Sem Wo',
    'Artigo (SKU)': o.item,
    'Descrição': o.desc,
    'Classe': o.classe,
    'Demanda': o.demanda || '-',
    'Cliente / Solicitante': o.cliente || '-',
    'Datas Programadas': o.datas.map(d => `${formatDateBR(d.dataIso)} (${d.qtd})`).join(' | '),
    'Programado no Plano (Cx)': o.progTotal,
    'Quantidade da Ordem SAP (GMEIN)': o.qtdOrdemAlvo,
    'Dúzias Previstas': Math.round(o.duziasTotal),
    'Horas Programadas': Number(o.horasTotal.toFixed(2)),
    'Realizado SAP (GMEIN)': o.realizado,
    'Saldo Pendente Ordem (Cx)': o.saldo,
    '% Concluído Ordem': `${o.pct}%`,
    'Observação PCP': o.notaPcp || ''
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  const nomeAba = tipo === 'finalizadas' ? 'Ordens_Finalizadas' : 'Ordens_Pendentes';
  XLSX.utils.book_append_sheet(wb, ws, nomeAba);

  const dataArquivo = refDateInput.value || new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `PCP_MTG_FAB_${nomeAba}_${dataArquivo}.xlsx`);
}

initCloudSync();