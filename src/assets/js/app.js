/* ═══════════════════════════════════════════════════════════════════
   TAF-EB — App principal (renderer)
   Portaria EME/C Ex Nº 850 / 2022
   ═══════════════════════════════════════════════════════════════════ */

'use strict';

// ── Estado global ────────────────────────────────────────────────────────
const state = {
  paginaAtual: 'dashboard',
  avaliacoes: [],
  militares: [],
  avaliacaoSelecionada: null,
  militarSelecionado: null,
};

// ── Dados de referência ───────────────────────────────────────────────────
const POSTOS_GRADUACOES = [
  // Oficiais Generais
  'Gen Ex', 'Gen Div', 'Gen Bda',
  // Oficiais Superiores
  'Cel', 'TC', 'Maj',
  // Oficiais Intermediários / Subalternos
  'Cap', '1º Ten', '2º Ten', 'Asp Off',
  // Praças
  'Sub Ten', '1º Sgt', '2º Sgt', '3º Sgt', 'Cb', 'Sd',
  // Ensino
  'Cadete', 'Aluno',
];

// ── Armas, Quadros e Serviços do EB (eb.mil.br/o-exercito/armas-quadros-e-servicos)
// Mapeamento: arma → LEM conforme Portaria EB20-D-03.053
const ARMAS = [
  // ── LEMB — Linha de Ensino Militar Bélico ────────────────────────────
  { value: 'INF',   label: 'Infantaria',              sigla: 'INF',  lem: 'LEMB', grupo: 'Armas' },
  { value: 'CAV',   label: 'Cavalaria',               sigla: 'CAV',  lem: 'LEMB', grupo: 'Armas' },
  { value: 'ART',   label: 'Artilharia',              sigla: 'ART',  lem: 'LEMB', grupo: 'Armas' },
  { value: 'ENG',   label: 'Engenharia',              sigla: 'ENG',  lem: 'LEMB', grupo: 'Armas' },
  { value: 'COM',   label: 'Comunicações',            sigla: 'COM',  lem: 'LEMB', grupo: 'Armas' },
  { value: 'AvEx',  label: 'Aviação do Exército',     sigla: 'AvEx', lem: 'LEMB', grupo: 'Armas' },
  { value: 'TOP',   label: 'Topografia',              sigla: 'TOP',  lem: 'LEMB', grupo: 'Armas' },
  // ── LEMS — Linha de Ensino Militar de Saúde ──────────────────────────
  { value: 'SSau',  label: 'Serviço de Saúde',        sigla: 'SSau', lem: 'LEMS', grupo: 'Serviços de Saúde' },
  // ── LEMC — Linha de Ensino Militar Complementar ───────────────────────
  { value: 'MatBel',label: 'Material Bélico',         sigla: 'MatBel',lem: 'LEMC', grupo: 'Quadros e Serviços' },
  { value: 'Int',   label: 'Intendência',             sigla: 'Int',  lem: 'LEMC', grupo: 'Quadros e Serviços' },
  { value: 'QEM',   label: 'Quadro de Engenheiros Militares', sigla: 'QEM', lem: 'LEMC', grupo: 'Quadros e Serviços' },
  { value: 'QAO',   label: 'Quadro Auxiliar de Oficiais',     sigla: 'QAO', lem: 'LEMC', grupo: 'Quadros e Serviços', obs: 'excluído PPM mesmo em Op' },
  { value: 'QCO',   label: 'Quadro Complementar de Oficiais', sigla: 'QCO', lem: 'LEMC', grupo: 'Quadros e Serviços' },
  { value: 'SAREX', label: 'Serviço de Assistência Religiosa',sigla: 'SAREX',lem: 'LEMC', grupo: 'Quadros e Serviços' },
  { value: 'MUS',   label: 'Músicos',                 sigla: 'MUS',  lem: 'LEMC', grupo: 'Quadros e Serviços', obs: 'PBD fixo per Portaria §2.3.3' },
  // ── LEMCT — Linha de Ensino Militar Científico-Tecnológico ───────────
  { value: 'LEMCT', label: 'Científico-Tecnológico (IME/ITA)',sigla:'LEMCT',lem: 'LEMCT', grupo: 'Científico-Tecnológico' },
];

/** Retorna o objeto arma a partir do value */
function getArma(value) {
  return ARMAS.find(a => a.value === value) || null;
}

/** Dado a arma e a situação funcional/tipo de OM, retorna o padrão exigido */
function calcularPadrao(armaValue, situacaoFuncional) {
  const arma = getArma(armaValue);
  if (!arma) return 'PBD';
  const lem = arma.lem;
  const isLEMB = lem === 'LEMB';

  if (situacaoFuncional === 'OM_F_Emp_Estrt' && isLEMB) return 'PED';
  if (situacaoFuncional === 'OM_Op'          && isLEMB) return 'PAD';
  // LEMS/LEMC/LEMCT: máximo PAD mesmo em Op (Portaria §2.4 e §2.5.3)
  if (situacaoFuncional === 'OM_Op' && !isLEMB) return 'PAD';
  return 'PBD';
}

// Grupos de arma para o <select> com <optgroup>
const GRUPOS_ARMA = [...new Set(ARMAS.map(a => a.grupo))];

function selectArmas(selected) {
  return GRUPOS_ARMA.map(g => `
    <optgroup label="${g}">
      ${ARMAS.filter(a => a.grupo === g).map(a =>
        `<option value="${a.value}" ${selected === a.value ? 'selected' : ''}>${a.sigla} — ${a.label}</option>`
      ).join('')}
    </optgroup>`).join('');
}

const SITUACOES_FUNCIONAIS = [
  { value: 'OM_NOp',         label: 'OM Não Operativa (NOp)' },
  { value: 'OM_Op',          label: 'OM Operativa (Op)' },
  { value: 'OM_F_Emp_Estrt', label: 'OM F Emp Estrt / Op Especiais' },
  { value: 'Estb_Ens',       label: 'Estabelecimento de Ensino' },
];

const TIPOS_TAF = [
  { value: 'diagnostica', label: 'Avaliação Diagnóstica' },
  { value: '1taf',        label: '1º TAF' },
  { value: '2taf',        label: '2º TAF' },
  { value: '3taf',        label: '3º TAF' },
];

// ── Bootstrap ─────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', async () => {
  const caminho = await window.api.db.caminho();
  document.getElementById('dbPath').textContent = caminho;
  await carregarDados();
  navegarPara('dashboard');
});

async function carregarDados() {
  state.avaliacoes = await window.api.avaliacoes.listar();
  state.militares  = await window.api.militares.listar();
}

// ── Navegação ─────────────────────────────────────────────────────────────
function navegarPara(pagina, btnEl) {
  state.paginaAtual = pagina;
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  if (btnEl) {
    btnEl.classList.add('active');
  } else {
    const btn = document.querySelector(`[data-page="${pagina}"]`);
    if (btn) btn.classList.add('active');
  }
  renderizarPagina(pagina);
}

async function renderizarPagina(pagina) {
  await carregarDados();
  const area = document.getElementById('contentArea');
  switch (pagina) {
    case 'dashboard':   area.innerHTML = paginaDashboard(); break;
    case 'avaliacoes':  area.innerHTML = paginaAvaliacoes(); break;
    case 'lancar': {
      const pid = _tafAvalPendente;
      _tafAvalPendente = null;
      area.innerHTML = paginaLancar(pid);
      // Se veio com avalId pendente, carregar automaticamente
      if (pid) setTimeout(abrirTabelaTAF, 100);
      break;
    }
    case 'militares':   area.innerHTML = paginaMilitares(); break;
    case 'ficha':       area.innerHTML = paginaFicha(); break;
    case 'importar':    area.innerHTML = paginaImportar(); break;
    default:            area.innerHTML = '<p>Página não encontrada.</p>';
  }
}

// ═════════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═════════════════════════════════════════════════════════════════════════
function paginaDashboard() {
  const totalMil = state.militares.length;
  const totalAval = state.avaliacoes.length;
  return `
  <div class="page-header">
    <div>
      <div class="page-title">PAINEL GERAL</div>
      <div class="page-subtitle">Visão geral do sistema TAF-EB · Portaria EME Nº 850/2022</div>
    </div>
  </div>

  <div class="stats-grid">
    <div class="stat-card">
      <div class="stat-val">${totalMil}</div>
      <div class="stat-label">Militares cadastrados</div>
    </div>
    <div class="stat-card">
      <div class="stat-val">${totalAval}</div>
      <div class="stat-label">TAFs cadastrados</div>
    </div>
    <div class="stat-card">
      <div class="stat-val">${new Date().getFullYear()}</div>
      <div class="stat-label">Ano corrente</div>
    </div>
  </div>

  <div class="card">
    <div class="card-title">⚡ AÇÕES RÁPIDAS</div>
    <div style="display:flex;gap:10px;flex-wrap:wrap;">
      <button class="btn btn-primary" onclick="navegarPara('avaliacoes',null)">+ Novo TAF</button>
      <button class="btn btn-ouro" onclick="navegarPara('militares',null)">+ Novo Militar</button>
      <button class="btn btn-ghost" onclick="navegarPara('lancar',null)">✎ Lançar Resultados</button>
      <button class="btn btn-ghost" onclick="navegarPara('avaliacoes',null)">⊞ Ver Resultados</button>
    </div>
  </div>

  <div class="card">
    <div class="card-title">📋 AVALIAÇÕES ATIVAS</div>
    ${state.avaliacoes.length === 0 ? '<div class="empty-state"><div class="empty-icon">📋</div><p>Nenhum TAF cadastrado ainda.</p></div>' : `
    <div class="table-wrap">
      <table>
        <thead><tr>
          <th>Descrição</th><th>Tipo</th><th>Ano</th>
          <th class="center">1ª Chamada</th><th class="center">2ª Chamada</th><th class="center">3ª Chamada</th>
          <th class="center">Ações</th>
        </tr></thead>
        <tbody>
          ${state.avaliacoes.map(a => `
          <tr>
            <td><strong>${a.descricao}</strong></td>
            <td>${labelTipo(a.tipo)}</td>
            <td>${a.ano}</td>
            <td class="center">${formatarData(a.data_1_chamada)}</td>
            <td class="center">${formatarData(a.data_2_chamada)}</td>
            <td class="center">${formatarData(a.data_3_chamada)}</td>
            <td class="center">
              <button class="btn btn-ghost btn-sm" onclick="navegarComAval('lancar', ${a.id})">Lançar</button>
            </td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>`}
  </div>`;
}

// ═════════════════════════════════════════════════════════════════════════
// AVALIAÇÕES (TAFs) — Gerenciar + Resultados integrados
// ═════════════════════════════════════════════════════════════════════════
function paginaAvaliacoes() {
  return `
  <div class="page-header">
    <div>
      <div class="page-title">GERENCIAR AVALIAÇÕES TAF</div>
      <div class="page-subtitle">Clique em uma avaliação para ver resultados e gráficos</div>
    </div>
    <div class="page-actions">
      <button class="btn btn-ouro" onclick="abrirModalAvaliacao()">+ Nova Avaliação</button>
    </div>
  </div>

  ${state.avaliacoes.length === 0 ? `
  <div class="empty-state">
    <div class="empty-icon">📋</div>
    <p>Nenhuma avaliação cadastrada.<br>Clique em <strong>+ Nova Avaliação</strong> para começar.</p>
  </div>` : `
  <div class="card" style="padding:0;overflow:hidden">
    <table>
      <thead><tr>
        <th>Descrição</th><th>Tipo</th><th class="center">Ano</th>
        <th class="center">Padrão</th><th class="center">Tipo OM</th>
        <th class="center">1ª Chamada</th><th class="center">2ª</th>
        <th class="center">3ª</th><th class="center">Extra</th>
        <th class="center">Ações</th>
      </tr></thead>
      <tbody>
        ${state.avaliacoes.map(a => {
          const coresPad = { PBD:'badge-B', PAD:'badge-MB', PED:'badge-E' };
          const padrao = a.padrao || 'PAD';
          return `
        <tr style="cursor:pointer" onclick="abrirResultadosTAF(${a.id})" title="Clique para ver resultados">
          <td><strong>${a.descricao}</strong></td>
          <td>${labelTipo(a.tipo)}</td>
          <td class="center">${a.ano}</td>
          <td class="center"><span class="badge ${coresPad[padrao]||'badge-B'}" style="font-size:10px">${padrao}</span></td>
          <td class="center" style="font-size:12px">${labelSitFuncionalCurto(a.situacao_funcional)}</td>
          <td class="center">${formatarData(a.data_1_chamada)}</td>
          <td class="center">${formatarData(a.data_2_chamada)}</td>
          <td class="center">${formatarData(a.data_3_chamada)}</td>
          <td class="center">${formatarData(a.data_chamada_extra)}</td>
          <td class="center" style="white-space:nowrap" onclick="event.stopPropagation()">
            <button class="btn btn-ghost btn-sm" onclick="abrirModalAvaliacao(${a.id})">✎</button>
            <button class="btn btn-ghost btn-sm" onclick="navegarComAval('lancar',${a.id})" title="Lançar resultados">⊕</button>
            <button class="btn btn-danger btn-sm" onclick="excluirAvaliacao(${a.id})">✕</button>
          </td>
        </tr>`;}).join('')}
      </tbody>
    </table>
  </div>

  <!-- Painel de resultados do TAF selecionado -->
  <div id="painelResultadosTAF" style="display:none;margin-top:20px"></div>`}`;
}

function labelSitFuncionalCurto(sf) {
  return { OM_NOp:'NOp', OM_Op:'Op', OM_F_Emp_Estrt:'F Emp Estrt', Estb_Ens:'Estb Ens' }[sf] || sf || '—';
}

function abrirModalAvaliacao(id) {
  const aval = id ? state.avaliacoes.find(a => a.id === id) : null;
  const ano = new Date().getFullYear();

  abrirModal(aval ? 'Editar Avaliação TAF' : 'Nova Avaliação TAF', `
    <div class="form-grid">
      <div class="form-group form-full">
        <label>Descrição *</label>
        <input id="fd_desc" placeholder="Ex: 1º TAF 2025 – CAV" value="${aval?.descricao || ''}">
      </div>
      <div class="form-group">
        <label>Tipo *</label>
        <select id="fd_tipo">
          ${TIPOS_TAF.map(t => `<option value="${t.value}" ${aval?.tipo === t.value ? 'selected' : ''}>${t.label}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label>Ano *</label>
        <input type="number" id="fd_ano" value="${aval?.ano || ano}" min="2020" max="2040">
      </div>
      <div class="form-group">
        <label>Padrão exigido *</label>
        <select id="fd_padrao">
          <option value="PBD" ${(aval?.padrao||'PAD')==='PBD' ? 'selected':''}>PBD — Padrão Básico (mín. Regular)</option>
          <option value="PAD" ${(aval?.padrao||'PAD')==='PAD' ? 'selected':''}>PAD — Padrão Avançado (mín. Bom)</option>
          <option value="PED" ${aval?.padrao==='PED' ? 'selected':''}>PED — Padrão Especial (mín. Muito Bom)</option>
        </select>
      </div>
      <div class="form-group">
        <label>Tipo de OM</label>
        <select id="fd_sitfunc">
          ${SITUACOES_FUNCIONAIS.map(s => `<option value="${s.value}" ${aval?.situacao_funcional === s.value ? 'selected':''}>${s.label}</option>`).join('')}
        </select>
      </div>
    </div>
    <div class="secao-titulo">DATAS DAS CHAMADAS</div>
    <div class="form-grid">
      <div class="form-group">
        <label>1ª Chamada</label>
        <input type="date" id="fd_d1" value="${aval?.data_1_chamada || ''}">
      </div>
      <div class="form-group">
        <label>2ª Chamada <span class="text-muted">(≤30 dias após 1ª)</span></label>
        <input type="date" id="fd_d2" value="${aval?.data_2_chamada || ''}">
      </div>
      <div class="form-group">
        <label>3ª Chamada</label>
        <input type="date" id="fd_d3" value="${aval?.data_3_chamada || ''}">
      </div>
      <div class="form-group">
        <label>Chamada Extra</label>
        <input type="date" id="fd_dex" value="${aval?.data_chamada_extra || ''}">
      </div>
    </div>`,
    `<button class="btn btn-ghost" onclick="fecharModalForce()">Cancelar</button>
     <button class="btn btn-ouro" onclick="salvarAvaliacao(${id || 0})">💾 Salvar</button>`
  );
}

async function salvarAvaliacao(id) {
  const dados = {
    id: id || null,
    descricao:          document.getElementById('fd_desc').value.trim(),
    tipo:               document.getElementById('fd_tipo').value,
    ano:                parseInt(document.getElementById('fd_ano').value),
    padrao:             document.getElementById('fd_padrao').value,
    situacao_funcional: document.getElementById('fd_sitfunc').value,
    data_1_chamada:     document.getElementById('fd_d1').value || null,
    data_2_chamada:     document.getElementById('fd_d2').value || null,
    data_3_chamada:     document.getElementById('fd_d3').value || null,
    data_chamada_extra: document.getElementById('fd_dex').value || null,
  };
  if (!dados.descricao) { toast('Informe a descrição.', true); return; }
  await window.api.avaliacoes.salvar(dados);
  fecharModal();
  toast('Avaliação salva com sucesso!');
  await renderizarPagina('avaliacoes');
}

// Recalcula todas as menções de uma avaliação usando a data correta da chamada
async function recalcularMencoes(avalId) {
  const aval = state.avaliacoes.find(a => a.id === avalId);
  if (!confirm(`Recalcular todas as menções de "${aval.descricao}" usando a data correta de cada chamada?\n\nIsso corrige resultados salvos com idade errada (data de hoje em vez da data do TAF).`)) return;

  const lista = await window.api.resultados.listar(avalId);
  if (!lista.length) { toast('Nenhum resultado para recalcular.', true); return; }

  toast(`Recalculando ${lista.length} resultados...`);

  let cnt = 0;
  for (const r of lista) {
    // Data correta da chamada
    const dataChamada = r.chamada === 1 ? aval.data_1_chamada
                      : r.chamada === 2 ? aval.data_2_chamada
                      : r.chamada === 3 ? aval.data_3_chamada
                      : aval.data_chamada_extra;

    const fallbackData = aval.ano ? `${aval.ano}-01-01` : null;
    const dataRef = dataChamada || fallbackData;

    const armaObj = getArma(r.arma || '');
    const lem = armaObj?.lem || r.lem || 'LEMB';
    const ppmSeg = TAF.mmssParaSegundos(r.ppm_tempo);

    const calc = TAF.calcularTAF({
      dataNascimento:    r.data_nascimento,
      dataChamada:       dataRef,
      anoTAF:            aval.ano || null,
      lem,
      sexo:              r.sexo,
      situacaoFuncional: r.situacao_funcional,
      padrao:            aval.padrao || null,
      postoGraduacao:    r.posto_graduacao,
      corrida:           r.corrida_12min,
      flexao:            r.flexao_bracos,
      abdominal:         r.abdominal_supra,
      barra:             r.flexao_barra,
      ppmTempo:          ppmSeg,
      tafAlternativo:    false,
    });

    await window.api.resultados.salvar({
      militar_id:             r.militar_id,
      avaliacao_id:           avalId,
      chamada:                r.chamada,
      corrida_12min:          r.corrida_12min,
      flexao_bracos:          r.flexao_bracos,
      abdominal_supra:        r.abdominal_supra,
      flexao_barra:           r.flexao_barra,
      ppm_tempo:              r.ppm_tempo,
      conceito_corrida:       calc.conceitos?.corrida       || null,
      conceito_flexao_bracos: calc.conceitos?.flexao        || null,
      conceito_abdominal:     calc.conceitos?.abdominal     || null,
      conceito_barra:         calc.conceitos?.barra         || null,
      conceito_ppm:           calc.suficiencias?.ppm        || null,
      conceito_global:        calc.conceitoGlobal,
      suficiencia:            calc.suficienciaGlobal,
      padrao_verificado:      calc.padrao,
      taf_alternativo: 0,
      observacoes: r.observacoes,
    });
    cnt++;
  }

  await carregarDados();
  toast(`✔ ${cnt} resultado(s) recalculados com data correta!`);
  reaplicarFiltrosResultados(avalId);
}

async function excluirAvaliacao(id) {
  if (!confirm('Excluir esta avaliação? Os resultados vinculados serão mantidos.')) return;
  await window.api.avaliacoes.excluir(id);
  toast('Avaliação removida.');
  await renderizarPagina('avaliacoes');
}

// ─── Resultados inline do TAF clicado ────────────────────────────────────
let _avalAtualId = null;

async function abrirResultadosTAF(avalId) {
  if (_avalAtualId === avalId) {
    _avalAtualId = null;
    document.getElementById('painelResultadosTAF').style.display = 'none';
    return;
  }
  _avalAtualId = avalId;
  // Resetar filtros ao abrir novo TAF
  _resFilters = { chamadas: new Set(), suficiencia: 'todos', nome: '', sexo: '', posto: '' };

  const painel = document.getElementById('painelResultadosTAF');
  painel.style.display = 'block';
  painel.innerHTML = '<div class="card"><p class="text-muted">Carregando...</p></div>';

  await reaplicarFiltrosResultados(avalId);
}

function renderResultadosTAF(painel, aval, lista) {
  if (!lista.length) {
    painel.innerHTML = `
      <div class="card">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px">
          <div class="page-title" style="font-size:18px">${aval.descricao}</div>
          <button class="btn btn-ghost btn-sm" onclick="abrirResultadosTAF(${aval.id})">✕ Fechar</button>
        </div>
        <div class="empty-state"><div class="empty-icon">📊</div><p>Nenhum resultado lançado.</p>
          <button class="btn btn-primary" style="margin-top:8px" onclick="navegarComAval('lancar',${aval.id})">+ Lançar Resultados</button>
        </div>
      </div>`;
    return;
  }

  // Filtros ativos
  const chamadaFiltro = document.getElementById('res_chamada_filtro')?.value || '';
  const sufFiltro     = document.getElementById('res_suf_filtro')?.value     || '';
  const nomeFiltro    = (document.getElementById('res_nome_filtro')?.value   || '').toLowerCase();

  let filtrada = lista;
  if (chamadaFiltro) filtrada = filtrada.filter(r => String(r.chamada) === chamadaFiltro);
  if (sufFiltro)     filtrada = filtrada.filter(r => r.suficiencia === sufFiltro);
  if (nomeFiltro)    filtrada = filtrada.filter(r =>
    (r.nome_guerra||r.nome||'').toLowerCase().includes(nomeFiltro));

  // Estatísticas
  const total = filtrada.length;
  const sufS  = filtrada.filter(r => r.suficiencia === 'S').length;
  const sufNS = filtrada.filter(r => r.suficiencia === 'NS').length;
  const sufNR = filtrada.filter(r => r.suficiencia === 'NR' || !r.suficiencia).length;

  // Contagem por menção global
  const ORDEM_M = ['E','MB','B','R','I'];
  const COR_M   = { E:'#1565C0', MB:'#2e7d32', B:'#f9a825', R:'#e65100', I:'#c62828' };
  const contMencao = {};
  ORDEM_M.forEach(m => contMencao[m] = filtrada.filter(r => r.conceito_global === m).length);

  // Contagem por menção de cada OII
  const OIIs = [
    { campo:'conceito_corrida',         label:'Corrida' },
    { campo:'conceito_flexao_bracos',   label:'Flexão' },
    { campo:'conceito_abdominal',       label:'Abdominal' },
    { campo:'conceito_barra',           label:'Barra' },
  ];
  const contOII = {};
  OIIs.forEach(o => {
    contOII[o.campo] = {};
    ORDEM_M.forEach(m => contOII[o.campo][m] = filtrada.filter(r => r[o.campo] === m).length);
  });

  painel.innerHTML = `
  <div class="card" style="padding:0;overflow:hidden">
    <!-- Cabeçalho do painel -->
    <div style="background:var(--verde-eb);padding:12px 16px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">
      <div>
        <span style="color:var(--ouro-claro);font-family:'Barlow Condensed',sans-serif;font-size:20px;letter-spacing:1px;font-weight:700">${aval.descricao}</span>
        <span style="color:rgba(255,255,255,.5);font-size:12px;margin-left:10px">${labelTipo(aval.tipo)} · ${aval.ano}</span>
      </div>
      <div style="display:flex;gap:8px;align-items:center">
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="exportarPDF(${aval.id})">📄 PDF Geral</button>
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="exportarWord(${aval.id})">📝 Word Geral</button>
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="exportarFichasLotePDF(${aval.id})">👤📄 Fichas PDF</button>
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="exportarFichasLoteWord(${aval.id})">👤📝 Fichas Word</button>
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="abrirResultadosTAF(${aval.id})">✕</button>
      </div>
    </div>

    <!-- Filtros -->
    <div class="filtro-bar" style="padding:10px 16px;background:#f8f7f5;border-bottom:1px solid var(--bege-escuro)">
      <select id="res_chamada_filtro" onchange="reaplicarFiltrosResultados(${aval.id})">
        <option value="">Todas chamadas</option>
        <option value="1">1ª Chamada</option>
        <option value="2">2ª Chamada</option>
        <option value="3">3ª Chamada</option>
        <option value="4">Extra</option>
      </select>
      <select id="res_suf_filtro" onchange="reaplicarFiltrosResultados(${aval.id})">
        <option value="">S + NS + NR</option>
        <option value="S">Apenas Suficientes</option>
        <option value="NS">Apenas NS</option>
        <option value="NR">Apenas NR</option>
      </select>
      <input id="res_nome_filtro" placeholder="🔍 Nome de guerra..." oninput="reaplicarFiltrosResultados(${aval.id})" style="min-width:180px">
      <button class="btn btn-ghost btn-sm" onclick="navegarComAval('lancar',${aval.id})">⊕ Lançar / Editar</button>
    </div>

    <div style="padding:16px">

      <!-- Cards de resumo -->
      <div class="stats-grid" style="margin-bottom:16px">
        <div class="stat-card"><div class="stat-val">${total}</div><div class="stat-label">Avaliados</div></div>
        <div class="stat-card"><div class="stat-val" style="color:#7ddb8a">${sufS}</div><div class="stat-label">Suficientes</div></div>
        <div class="stat-card"><div class="stat-val" style="color:#e87070">${sufNS}</div><div class="stat-label">Não Suficientes</div></div>
        <div class="stat-card"><div class="stat-val" style="color:rgba(255,255,255,.5)">${sufNR}</div><div class="stat-label">NR</div></div>
        <div class="stat-card"><div class="stat-val" style="color:#7ddb8a">${total>0?(sufS/total*100).toFixed(1):0}%</div><div class="stat-label">% Suficientes</div></div>
      </div>

      <!-- Gráficos: Suficiência + Menção Global -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px">
        <div class="card" style="margin-bottom:0">
          <div class="card-title">Quantidade / Suficiência</div>
          <canvas id="grafSufQtd" height="180"></canvas>
        </div>
        <div class="card" style="margin-bottom:0">
          <div class="card-title">Percentual / Suficiência</div>
          <canvas id="grafSufPct" height="180"></canvas>
        </div>
        <div class="card" style="margin-bottom:0">
          <div class="card-title">Quantidade / Menção Global</div>
          <canvas id="grafMenQtd" height="180"></canvas>
        </div>
        <div class="card" style="margin-bottom:0">
          <div class="card-title">Percentual / Menção Global</div>
          <canvas id="grafMenPct" height="180"></canvas>
        </div>
      </div>

      <!-- Gráficos por OII -->
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:16px;margin-bottom:16px">
        ${OIIs.map(o => `
        <div class="card" style="margin-bottom:0">
          <div class="card-title">${o.label} — Menção</div>
          <canvas id="grafOII_${o.campo}" height="160"></canvas>
        </div>`).join('')}
      </div>

      <!-- Tabela de resultados -->
      <div class="card" style="padding:0;overflow:hidden;margin-bottom:0">
        <div style="padding:8px 14px;background:var(--bege-escuro);font-family:'Barlow Condensed',sans-serif;font-size:13px;letter-spacing:.5px;color:var(--verde-eb)">
          LISTA DE RESULTADOS — ${filtrada.length} militar(es)
        </div>
        <div class="table-wrap">
          <table style="font-size:12px">
            <thead><tr>
              <th>Nome</th><th>Posto</th><th class="center">Sx</th>
              <th class="center">Corrida</th><th class="center">Flexão</th>
              <th class="center">Abdom</th><th class="center">Barra</th>
              <th class="center">PPM</th><th class="center">Global</th>
              <th class="center">Sufic.</th><th class="center">Cham.</th>
            </tr></thead>
            <tbody>
              ${filtrada.map(r => `
              <tr>
                <td><strong>${r.nome_guerra||r.nome}</strong></td>
                <td>${r.posto_graduacao}</td>
                <td class="center">${r.sexo}</td>
                <td class="center">${r.corrida_12min||'—'} ${badgeConceito(r.conceito_corrida)}</td>
                <td class="center">${r.flexao_bracos||'—'} ${badgeConceito(r.conceito_flexao_bracos)}</td>
                <td class="center">${r.abdominal_supra||'—'} ${badgeConceito(r.conceito_abdominal)}</td>
                <td class="center">${r.flexao_barra||'—'} ${badgeConceito(r.conceito_barra)}</td>
                <td class="center">${r.ppm_tempo||'—'} ${badgeConceito(r.conceito_ppm)}</td>
                <td class="center">${badgeConceito(r.conceito_global)}</td>
                <td class="center">${badgeSuficiencia(r.suficiencia)}</td>
                <td class="center">${r.chamada}ª</td>
              </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>`;

  // Renderizar gráficos com Chart.js via canvas inline
  setTimeout(() => {
    renderGraficos(sufS, sufNS, sufNR, contMencao, contOII, OIIs, COR_M, ORDEM_M);
  }, 50);
}

// ── Funções de canvas para gráficos no painel ─────────────────────────────
function renderGraficos(sufS, sufNS, sufNR, contMencao, contOII, OIIs, COR_M, ORDEM_M) {
  const total = sufS + sufNS + sufNR;
  desenharBarras('grafSufQtd',['S','NS','NR'],[sufS,sufNS,sufNR],['#2e7d32','#c62828','#9e9e9e']);
  desenharPizza ('grafSufPct',['S','NS','NR'],[sufS,sufNS,sufNR],['#2e7d32','#c62828','#9e9e9e'],total);
  desenharBarras('grafMenQtd',ORDEM_M,ORDEM_M.map(m=>contMencao[m]),ORDEM_M.map(m=>COR_M[m]));
  const tm = ORDEM_M.reduce((s,m)=>s+contMencao[m],0);
  desenharPizza ('grafMenPct',ORDEM_M,ORDEM_M.map(m=>contMencao[m]),ORDEM_M.map(m=>COR_M[m]),tm);
  OIIs.forEach(o => {
    const vals = ORDEM_M.map(m=>contOII[o.campo][m]);
    const tot  = vals.reduce((s,v)=>s+v,0);
    desenharBarras('grafOII_'+o.campo, ORDEM_M, vals, ORDEM_M.map(m=>COR_M[m]));
    if (tot > 0) desenharPizza('grafOIIpct_'+o.campo, ORDEM_M, vals, ORDEM_M.map(m=>COR_M[m]), tot);
  });
}

function _canvasW(canvas) {
  // canvas.offsetWidth pode ser 0 se o elemento não foi pintado ainda
  return canvas.offsetWidth || canvas.clientWidth ||
         canvas.parentElement?.offsetWidth || canvas.parentElement?.clientWidth || 420;
}

function desenharBarras(canvasId, labels, valores, cores) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const W = canvas.width = _canvasW(canvas);
  const H = canvas.height || 160;
  const pad = { t:22, r:10, b:28, l:28 };
  const maxV = Math.max(...valores, 1);
  const n    = labels.length;
  const barW = Math.floor((W - pad.l - pad.r) / n * 0.58);
  const gap  = Math.floor((W - pad.l - pad.r) / n);

  ctx.clearRect(0, 0, W, H);
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);

  labels.forEach((lbl, i) => {
    const v  = valores[i] || 0;
    const bH = Math.max(Math.round((v / maxV) * (H - pad.t - pad.b)), v > 0 ? 2 : 0);
    const x  = pad.l + i * gap + (gap - barW) / 2;
    const y  = H - pad.b - bH;
    ctx.fillStyle = cores[i];
    if (ctx.roundRect) {
      ctx.beginPath(); ctx.roundRect(x, y, barW, bH || 1, 2); ctx.fill();
    } else {
      ctx.fillRect(x, y, barW, bH || 1);
    }
    if (v > 0) {
      ctx.fillStyle = '#333'; ctx.font = 'bold 11px sans-serif'; ctx.textAlign = 'center';
      ctx.fillText(v, x + barW/2, Math.max(y - 3, pad.t + 9));
    }
    ctx.fillStyle = '#555'; ctx.font = '11px sans-serif'; ctx.textAlign = 'center';
    ctx.fillText(lbl, x + barW/2, H - pad.b + 13);
  });
  ctx.strokeStyle = '#ddd'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(pad.l, H - pad.b); ctx.lineTo(W - pad.r, H - pad.b); ctx.stroke();
}

function desenharPizza(canvasId, labels, valores, cores, total) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const W = canvas.width = _canvasW(canvas);
  const H = canvas.height || 160;
  ctx.clearRect(0, 0, W, H); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
  if (!total) { ctx.fillStyle='#ccc'; ctx.font='12px sans-serif'; ctx.textAlign='center'; ctx.fillText('Sem dados',W/2,H/2); return; }

  const cx = W * 0.36, cy = H / 2, r = Math.min(cx - 4, cy - 8);
  let ang = -Math.PI / 2;
  valores.forEach((v, i) => {
    if (!v) return;
    const sl = (v / total) * 2 * Math.PI;
    ctx.beginPath(); ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, ang, ang + sl);
    ctx.closePath();
    ctx.fillStyle = cores[i]; ctx.fill();
    ctx.strokeStyle = '#fff'; ctx.lineWidth = 2; ctx.stroke();
    ang += sl;
  });
  const lx = cx + r + 8, ly0 = cy - (labels.length * 17) / 2;
  labels.forEach((lbl, i) => {
    if (!valores[i]) return;
    const pct = (valores[i] / total * 100).toFixed(1);
    const ly  = ly0 + i * 18;
    ctx.fillStyle = cores[i]; ctx.fillRect(lx, ly, 11, 11);
    ctx.fillStyle = '#333'; ctx.font = '10px sans-serif'; ctx.textAlign = 'left';
    ctx.fillText(`${lbl}  ${pct}%`, lx + 14, ly + 9);
  });
}


// ── Estado dos filtros de resultados ──────────────────────────────────────
let _resFilters = {
  chamadas: new Set(),      // Set vazio = todas
  suficiencia: 'todos',     // todos | S | NS | NR | cadastrados
  nome: '',
  sexo: '',
  posto: '',
};
let _resAvalId = null;
let _resListaCompleta = [];  // inclui militares sem resultado

async function reaplicarFiltrosResultados(avalId) {
  _resAvalId = avalId;
  const aval = state.avaliacoes.find(a => a.id === avalId);
  const painel = document.getElementById('painelResultadosTAF');
  if (!painel) return;

  // Carregar lista completa (militares com E sem resultado)
  const chamadaRef = _resFilters.chamadas.size === 1
    ? [..._resFilters.chamadas][0]
    : 1; // default 1ª para listarComMilitares
  _resListaCompleta = await window.api.resultados.listarComMilitares(avalId, chamadaRef);

  // Resultados com lançamento (chamadas selecionadas)
  let listaRes = await window.api.resultados.listar(avalId);

  // Filtrar por chamada(s)
  if (_resFilters.chamadas.size > 0) {
    listaRes = listaRes.filter(r => _resFilters.chamadas.has(String(r.chamada)));
  }

  renderResultadosTAF(painel, aval, listaRes, _resListaCompleta);
}

function renderResultadosTAF(painel, aval, lista, listaCompleta) {
  // ── Contagens ───────────────────────────────────────────────────────────
  const total    = lista.length;
  const sufS     = lista.filter(r => r.suficiencia === 'S').length;
  const sufNS    = lista.filter(r => r.suficiencia === 'NS').length;
  const sufNR    = lista.filter(r => r.suficiencia === 'NR' || r.conceito_global === 'NR').length;
  const cadastrados = listaCompleta?.length || total;
  const semResult   = listaCompleta?.filter(r => !r.resultado_id && !r.conceito_global).length || 0;

  const ORDEM_M = ['E','MB','B','R','I'];
  const COR_M   = { E:'#1565C0', MB:'#2e7d32', B:'#f9a825', R:'#e65100', I:'#c62828' };

  const contMencao = {};
  ORDEM_M.forEach(m => contMencao[m] = lista.filter(r => r.conceito_global === m).length);

  const OIIs = [
    { campo:'conceito_corrida',       label:'Corrida' },
    { campo:'conceito_flexao_bracos', label:'Flexão' },
    { campo:'conceito_abdominal',     label:'Abdominal' },
    { campo:'conceito_barra',         label:'Barra' },
  ];
  const contOII = {};
  OIIs.forEach(o => {
    contOII[o.campo] = {};
    ORDEM_M.forEach(m => contOII[o.campo][m] = lista.filter(r => r[o.campo] === m).length);
  });

  // ── Filtrar lista para exibição na tabela ────────────────────────────────
  let filtrada = lista;

  // Filtro suficiência
  if (_resFilters.suficiencia === 'S')  filtrada = filtrada.filter(r => r.suficiencia === 'S');
  if (_resFilters.suficiencia === 'NS') filtrada = filtrada.filter(r => r.suficiencia === 'NS');
  if (_resFilters.suficiencia === 'NR') filtrada = filtrada.filter(r => r.suficiencia === 'NR' || r.conceito_global === 'NR');
  if (_resFilters.suficiencia === 'cadastrados') {
    // Mostrar todos os militares cadastrados, com ou sem resultado
    filtrada = (listaCompleta || []).map(row => ({
      ...row,
      conceito_global: row.conceito_global || null,
      suficiencia:     row.suficiencia     || null,
      _semResultado:   !row.conceito_global,
    }));
  }

  // Filtros de nome, sexo, posto
  if (_resFilters.nome)  filtrada = filtrada.filter(r => (r.nome_guerra||r.nome||'').toLowerCase().includes(_resFilters.nome));
  if (_resFilters.sexo)  filtrada = filtrada.filter(r => r.sexo === _resFilters.sexo);
  if (_resFilters.posto) filtrada = filtrada.filter(r => (r.posto_graduacao||'') === _resFilters.posto);

  // ── Postos únicos para checkbox ─────────────────────────────────────────
  const postosUnicos = [...new Set(state.militares.map(m => m.posto_graduacao))].sort();
  const chamadas     = ['1','2','3','4'];

  // ── HTML ─────────────────────────────────────────────────────────────────
  painel.innerHTML = `
  <div class="card" style="padding:0;overflow:hidden">

    <!-- Cabeçalho verde -->
    <div style="background:var(--verde-eb);padding:10px 14px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">
      <div>
        <span style="color:var(--ouro-claro);font-family:'Barlow Condensed',sans-serif;font-size:18px;font-weight:700;letter-spacing:.5px">${aval.descricao}</span>
        <span style="color:rgba(255,255,255,.45);font-size:12px;margin-left:8px">${labelTipo(aval.tipo)} · ${aval.ano}</span>
      </div>
      <div style="display:flex;gap:6px;flex-wrap:wrap">
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="exportarPDF(${aval.id})">📄 PDF</button>
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="exportarWord(${aval.id})">📝 Word</button>
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="exportarFichasLotePDF(${aval.id})">👤📄 Fichas PDF</button>
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="exportarFichasLoteWord(${aval.id})">👤📝 Fichas Word</button>
        <button class="btn btn-ghost btn-sm" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)" onclick="abrirResultadosTAF(${aval.id})">✕</button>
      </div>
    </div>

    <div style="padding:12px 14px">

      <!-- Cards de resumo -->
      <div class="stats-grid" style="margin-bottom:12px">
        <div class="stat-card"><div class="stat-val">${cadastrados}</div><div class="stat-label">Cadastrados</div></div>
        <div class="stat-card"><div class="stat-val" style="color:#7ddb8a">${sufS}</div><div class="stat-label">Suficientes</div></div>
        <div class="stat-card"><div class="stat-val" style="color:#e87070">${sufNS}</div><div class="stat-label">NS</div></div>
        <div class="stat-card"><div class="stat-val" style="color:#e0a800">${sufNR}</div><div class="stat-label">NR</div></div>
        <div class="stat-card"><div class="stat-val" style="color:rgba(255,255,255,.4)">${semResult}</div><div class="stat-label">Sem lançamento</div></div>
        <div class="stat-card"><div class="stat-val" style="color:#7ddb8a">${total>0?(sufS/total*100).toFixed(1):0}%</div><div class="stat-label">% Suficientes</div></div>
      </div>

      <!-- Gráficos -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
        <div class="card" style="margin-bottom:0"><div class="card-title" style="font-size:12px">Qtd / Suficiência</div><canvas id="grafSufQtd" height="160"></canvas></div>
        <div class="card" style="margin-bottom:0"><div class="card-title" style="font-size:12px">% / Suficiência</div><canvas id="grafSufPct" height="160"></canvas></div>
        <div class="card" style="margin-bottom:0"><div class="card-title" style="font-size:12px">Qtd / Menção Global</div><canvas id="grafMenQtd" height="160"></canvas></div>
        <div class="card" style="margin-bottom:0"><div class="card-title" style="font-size:12px">% / Menção Global</div><canvas id="grafMenPct" height="160"></canvas></div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px">
        ${OIIs.map(o=>`
        <div class="card" style="margin-bottom:0">
          <div class="card-title" style="font-size:11px">${o.label}</div>
          <canvas id="grafOII_${o.campo}" height="130"></canvas>
          <div style="margin-top:6px;border-top:1px solid var(--bege-escuro);padding-top:6px">
            <canvas id="grafOIIpct_${o.campo}" height="110"></canvas>
          </div>
        </div>`).join('')}
      </div>

      <!-- FILTROS ─────────────────────────────────────────────────────── -->
      <div class="card" style="margin-bottom:12px;padding:10px 14px">
        <div style="display:flex;flex-wrap:wrap;gap:14px;align-items:flex-start">

          <!-- Chamada: multi-select botão toggle -->
          <div>
            <div style="font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--texto-leve);margin-bottom:4px">Chamada</div>
            <div style="display:flex;gap:4px">
              ${chamadas.map(ch => `
              <button id="btn_ch_${ch}" onclick="toggleChamada('${ch}',${aval.id})"
                class="btn btn-sm ${_resFilters.chamadas.has(ch)?'btn-ouro':'btn-ghost'}"
                style="min-width:36px;padding:3px 8px;font-size:12px">${ch}ª</button>`).join('')}
              <button onclick="toggleChamada('all',${aval.id})"
                class="btn btn-sm ${_resFilters.chamadas.size===0?'btn-ouro':'btn-ghost'}"
                style="min-width:40px;padding:3px 8px;font-size:12px">Todas</button>
            </div>
          </div>

          <!-- Suficiência: radio select -->
          <div>
            <div style="font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--texto-leve);margin-bottom:4px">Exibir</div>
            <div style="display:flex;flex-wrap:wrap;gap:4px">
              ${[
                ['todos','Todos com resultado'],
                ['S','Suficientes'],
                ['NS','Não Suficientes'],
                ['NR','NR'],
                ['cadastrados','Todos cadastrados'],
              ].map(([val,lbl]) => `
              <label style="display:flex;align-items:center;gap:4px;font-size:12px;cursor:pointer;padding:3px 8px;border-radius:4px;border:1px solid ${_resFilters.suficiencia===val?'var(--ouro)':'var(--bege-escuro)'};background:${_resFilters.suficiencia===val?'rgba(201,162,39,.15)':'transparent'}">
                <input type="radio" name="res_suf_${aval.id}" value="${val}" ${_resFilters.suficiencia===val?'checked':''}
                  onchange="_resFilters.suficiencia=this.value;reaplicarFiltrosResultados(${aval.id})">
                ${lbl}
              </label>`).join('')}
            </div>
          </div>

          <!-- Nome -->
          <div>
            <div style="font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--texto-leve);margin-bottom:4px">Nome</div>
            <input id="res_nome_f" placeholder="🔍 Nome de guerra..." value="${_resFilters.nome}"
              oninput="_resFilters.nome=this.value.toLowerCase();reaplicarFiltrosResultados(${aval.id})"
              style="height:28px;font-size:12px;min-width:160px">
          </div>

          <!-- Sexo checkbox -->
          <div>
            <div style="font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--texto-leve);margin-bottom:4px">Sexo</div>
            <div style="display:flex;gap:6px">
              ${[['','Todos'],['M','♂ Masc'],['F','♀ Fem']].map(([val,lbl])=>`
              <label style="display:flex;align-items:center;gap:3px;font-size:12px;cursor:pointer">
                <input type="radio" name="res_sx_${aval.id}" value="${val}" ${_resFilters.sexo===val?'checked':''}
                  onchange="_resFilters.sexo=this.value;reaplicarFiltrosResultados(${aval.id})">
                ${lbl}
              </label>`).join('')}
            </div>
          </div>

          <!-- Posto checkboxes -->
          <div>
            <div style="font-size:10px;text-transform:uppercase;letter-spacing:.8px;color:var(--texto-leve);margin-bottom:4px">Posto/Graduação</div>
            <div style="display:flex;flex-wrap:wrap;gap:4px">
              <label style="font-size:11px;cursor:pointer;display:flex;align-items:center;gap:3px">
                <input type="radio" name="res_posto_${aval.id}" value="" ${!_resFilters.posto?'checked':''}
                  onchange="_resFilters.posto=this.value;reaplicarFiltrosResultados(${aval.id})"> Todos
              </label>
              ${postosUnicos.map(p=>`
              <label style="font-size:11px;cursor:pointer;display:flex;align-items:center;gap:3px">
                <input type="radio" name="res_posto_${aval.id}" value="${p}" ${_resFilters.posto===p?'checked':''}
                  onchange="_resFilters.posto=this.value;reaplicarFiltrosResultados(${aval.id})"> ${p}
              </label>`).join('')}
            </div>
          </div>

          <!-- Botão lançar -->
          <div style="align-self:flex-end">
            <button class="btn btn-primary btn-sm" onclick="navegarComAval('lancar',${aval.id})">⊕ Lançar / Editar</button>
          <button class="btn btn-ghost btn-sm" title="Recalcular todas as menções com a data correta da chamada" onclick="recalcularMencoes(${aval.id})">🔄 Recalcular</button>
          </div>
        </div>
      </div>

      <!-- TABELA DE RESULTADOS -->
      <div style="font-size:11px;color:var(--texto-leve);margin-bottom:4px">${filtrada.length} militar(es) exibido(s)</div>
      <div class="table-wrap">
        <table style="font-size:11px">
          <thead><tr>
            <th>Nome</th><th>Posto</th><th class="center">Sx</th>
            <th class="center">Corrida</th><th class="center">Flexão</th>
            <th class="center">Abdom</th><th class="center">Barra</th>
            <th class="center">PPM</th><th class="center">Global</th>
            <th class="center">Sufic.</th><th class="center">Cham.</th>
            <th class="center">Ficha</th>
          </tr></thead>
          <tbody>
            ${filtrada.map(r => {
              const semRes = r._semResultado || (!r.conceito_global && !r.corrida_12min);
              return `<tr style="${semRes?'opacity:.6':''}">
                <td><strong>${r.nome_guerra||r.nome}</strong></td>
                <td>${r.posto_graduacao}</td>
                <td class="center">${r.sexo}</td>
                <td class="center">${r.corrida_12min||'—'} ${badgeConceito(r.conceito_corrida)}</td>
                <td class="center">${r.flexao_bracos||'—'} ${badgeConceito(r.conceito_flexao_bracos)}</td>
                <td class="center">${r.abdominal_supra||'—'} ${badgeConceito(r.conceito_abdominal)}</td>
                <td class="center">${r.flexao_barra||'—'} ${badgeConceito(r.conceito_barra)}</td>
                <td class="center">${r.ppm_tempo||'—'} ${badgeConceito(r.conceito_ppm)}</td>
                <td class="center">${semRes?'<span style="color:#aaa;font-size:10px">s/res</span>':badgeConceito(r.conceito_global)}</td>
                <td class="center">${semRes?'—':badgeSuficiencia(r.suficiencia)}</td>
                <td class="center">${r.chamada?r.chamada+'ª':'—'}</td>
                <td class="center">
                  <button class="btn btn-ghost btn-sm" style="padding:2px 6px"
                    onclick="navegarComMilitar('ficha',${r.militar_id||r.id})">📋</button>
                </td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>
    </div>
  </div>`;

  // Gráficos canvas — aguardar layout ser pintado antes de medir offsetWidth
  setTimeout(() => {
    const sufNR2 = lista.filter(r => r.suficiencia==='NR'||r.conceito_global==='NR').length;
    desenharBarras('grafSufQtd',['S','NS','NR'],[sufS,sufNS,sufNR2],['#2e7d32','#c62828','#9e9e9e']);
    desenharPizza ('grafSufPct',['S','NS','NR'],[sufS,sufNS,sufNR2],['#2e7d32','#c62828','#9e9e9e'],total);
    desenharBarras('grafMenQtd',ORDEM_M,ORDEM_M.map(m=>contMencao[m]),ORDEM_M.map(m=>COR_M[m]));
    desenharPizza ('grafMenPct',ORDEM_M,ORDEM_M.map(m=>contMencao[m]),ORDEM_M.map(m=>COR_M[m]),total);
    OIIs.forEach(o => {
      const vals  = ORDEM_M.map(m=>contOII[o.campo][m]);
      const total_oii = vals.reduce((s,v)=>s+v,0);
      desenharBarras('grafOII_'+o.campo, ORDEM_M, vals, ORDEM_M.map(m=>COR_M[m]));
      if (total_oii > 0)
        desenharPizza('grafOIIpct_'+o.campo, ORDEM_M, vals, ORDEM_M.map(m=>COR_M[m]), total_oii);
    });
  }, 200);
}

function toggleChamada(ch, avalId) {
  if (ch === 'all') {
    _resFilters.chamadas.clear();
  } else {
    if (_resFilters.chamadas.has(ch)) _resFilters.chamadas.delete(ch);
    else _resFilters.chamadas.add(ch);
  }
  reaplicarFiltrosResultados(avalId);
}



// ═════════════════════════════════════════════════════════════════════════
// GRÁFICOS NATIVOS OOXML (DrawingML Chart) — editáveis no Word
// ─────────────────────────────────────────────────────────────────────────
// Cada gráfico é um arquivo word/charts/chartN.xml dentro do .docx ZIP.
// O Word abre, exibe e permite editar tipo, cores, valores, rótulos etc.
// Sem imagens — 100% nativo Office.
// ═════════════════════════════════════════════════════════════════════════

// Cor por menção/suficiência (hex 6 dígitos, sem #)
const _CM = {E:'1565C0',MB:'2E7D32',B:'558B2F',R:'F9A825',I:'C62828',
             S:'2E7D32',NS:'C62828',NR:'9E9E9E'};
const _MVAL = {E:5,MB:4,B:3,R:2,I:1};
const _MLBL = ['','I','R','B','MB','E'];

// ── Barras (col) — dados diretos ──────────────────────────────────────────
function _ooxmlBarras(titulo, cats, vals, cores, chartId) {
  const n = cats.length;
  const ax1b = 2000 + (chartId||1)*2;
  const ax2b = 2001 + (chartId||1)*2;
  const ptsCat = cats.map((c,i)=>`<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('');
  const ptsVal = vals.map((v,i)=>`<c:pt idx="${i}"><c:v>${v||0}</c:v></c:pt>`).join('');
  const dpts   = cats.map((_,i)=>{
    const cor = cores?.[i] || '607D8B';
    return `<c:dPt><c:idx val="${i}"/>
      <c:spPr><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill>
        <a:ln><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill></a:ln>
      </c:spPr></c:dPt>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="pt-BR"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:rPr b="1" sz="1000"><a:solidFill><a:srgbClr val="2D4A1E"/></a:solidFill></a:rPr>
        <a:t>${titulo}</a:t></a:r></a:p></c:rich></c:tx>
      <c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea><c:layout/>
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          ${dpts}
          <c:dLbls>
            <c:spPr><a:noFill/></c:spPr>
            <c:txPr><a:bodyPr/><a:lstStyle/>
              <a:p><a:pPr><a:defRPr b="1" sz="900"/></a:pPr></a:p></c:txPr>
            <c:showLegendKey val="0"/><c:showVal val="1"/>
            <c:showCatName val="0"/><c:showSerName val="0"/>
            <c:showPercent val="0"/><c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:cat><c:strRef><c:f/>
            <c:strCache><c:ptCount val="${n}"/>${ptsCat}</c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef><c:f/>
            <c:numCache><c:formatCode>General</c:formatCode>
              <c:ptCount val="${n}"/>${ptsVal}
            </c:numCache>
          </c:numRef></c:val>
        </c:ser>
        <c:axId val="${ax1b}"/><c:axId val="${ax2b}"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="${ax1b}"/><c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:spPr><a:ln><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></c:spPr>
        <c:txPr><a:bodyPr/><a:lstStyle/>
          <a:p><a:pPr><a:defRPr b="1" sz="900"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:defRPr></a:pPr></a:p>
        </c:txPr>
        <c:crossAx val="${ax2b}"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="${ax2b}"/><c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:spPr><a:ln><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="${ax1b}"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="none"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

// ── Pizza — dados diretos ──────────────────────────────────────────────────
function _ooxmlPizza(titulo, cats, vals, cores) {
  const n = cats.length;
  const ax1b = 2000 + (chartId||1)*2;
  const ax2b = 2001 + (chartId||1)*2;
  const ptsCat = cats.map((c,i)=>`<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('');
  const ptsVal = vals.map((v,i)=>`<c:pt idx="${i}"><c:v>${v||0}</c:v></c:pt>`).join('');
  const dpts   = cats.map((_,i)=>{
    const cor = cores?.[i] || '607D8B';
    return `<c:dPt><c:idx val="${i}"/>
      <c:spPr><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill>
        <a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:w val="12700"/></a:ln>
      </c:spPr></c:dPt>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="pt-BR"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:rPr b="1" sz="1000"><a:solidFill><a:srgbClr val="2D4A1E"/></a:solidFill></a:rPr>
        <a:t>${titulo}</a:t></a:r></a:p></c:rich></c:tx>
      <c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea><c:layout/>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          ${dpts}
          <c:dLbls>
            <c:spPr><a:noFill/></c:spPr>
            <c:txPr><a:bodyPr/><a:lstStyle/>
              <a:p><a:pPr><a:defRPr b="1" sz="800"/></a:pPr></a:p></c:txPr>
            <c:showLegendKey val="0"/><c:showVal val="0"/>
            <c:showCatName val="1"/><c:showSerName val="0"/>
            <c:showPercent val="1"/><c:showBubbleSize val="0"/>
            <c:separator>: </c:separator>
          </c:dLbls>
          <c:cat><c:strRef><c:f/>
            <c:strCache><c:ptCount val="${n}"/>${ptsCat}</c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef><c:f/>
            <c:numCache><c:formatCode>General</c:formatCode>
              <c:ptCount val="${n}"/>${ptsVal}
            </c:numCache>
          </c:numRef></c:val>
        </c:ser>
        <c:firstSliceAng val="270"/>
      </c:pieChart>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/><c:overlay val="0"/>
      <c:txPr><a:bodyPr/><a:lstStyle/>
        <a:p><a:pPr><a:defRPr sz="900"/></a:pPr></a:p></c:txPr>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

// ── Linha de evolução (ficha individual) ──────────────────────────────────
function _ooxmlLinha(titulo, labels, mencoes, chartId) {
  const n = labels.length;
  const vals = mencoes.map(m => _MVAL[m] || 0);
  const ptsCat = labels.map((l,i)=>`<c:pt idx="${i}"><c:v>${l}</c:v></c:pt>`).join('');
  const ptsVal = vals.map((v,i)=>`<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
  // axIds únicos por chart — evita conflito quando há múltiplos charts no doc
  const ax1 = 1000 + (chartId||1)*2;
  const ax2 = 1001 + (chartId||1)*2;

  // Pontos coloridos por menção com rótulo
  const dpts = mencoes.map((m,i)=>{
    if (!m || !_MVAL[m]) return '';
    const cor = _CM[m] || '607D8B';
    return `<c:dPt><c:idx val="${i}"/>
      <c:marker>
        <c:symbol val="circle"/><c:size val="7"/>
        <c:spPr><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill>
          <a:ln><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill></a:ln>
        </c:spPr>
      </c:marker>
      <c:dLbl>
        <c:idx val="${i}"/>
        <c:spPr><a:noFill/></c:spPr>
        <c:txPr><a:bodyPr/><a:lstStyle/>
          <a:p><a:pPr><a:defRPr b="1" sz="900">
            <a:solidFill><a:srgbClr val="${cor}"/></a:solidFill>
          </a:defRPr></a:pPr></a:p>
        </c:txPr>
        <c:showLegendKey val="0"/><c:showVal val="0"/>
        <c:showCatName val="0"/><c:showSerName val="0"/>
        <c:showPercent val="0"/><c:showBubbleSize val="0"/>
        <c:tx><c:strRef><c:f/>
          <c:strCache><c:ptCount val="1"/>
            <c:pt idx="0"><c:v>${m}</c:v></c:pt>
          </c:strCache>
        </c:strRef></c:tx>
      </c:dLbl>
    </c:dPt>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:lang val="pt-BR"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:rPr b="1" sz="1000"><a:solidFill><a:srgbClr val="2D4A1E"/></a:solidFill></a:rPr>
        <a:t>${titulo}</a:t></a:r></a:p></c:rich></c:tx>
      <c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea><c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:spPr><a:noFill/>
            <a:ln w="25400"><a:solidFill><a:srgbClr val="2D4A1E"/></a:solidFill>
              <a:prstDash val="solid"/>
            </a:ln>
          </c:spPr>
          <c:marker><c:symbol val="circle"/><c:size val="5"/>
            <c:spPr><a:solidFill><a:srgbClr val="2D4A1E"/></a:solidFill></c:spPr>
          </c:marker>
          ${dpts}
          <c:dLbls>
            <c:showLegendKey val="0"/><c:showVal val="0"/>
            <c:showCatName val="0"/><c:showSerName val="0"/>
            <c:showPercent val="0"/><c:showBubbleSize val="0"/>
          </c:dLbls>
          <c:cat><c:strRef><c:f/>
            <c:strCache><c:ptCount val="${n}"/>${ptsCat}</c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef><c:f/>
            <c:numCache><c:formatCode>General</c:formatCode>
              <c:ptCount val="${n}"/>${ptsVal}
            </c:numCache>
          </c:numRef></c:val>
          <c:smooth val="0"/>
        </c:ser>
        <c:dLbls>
          <c:showLegendKey val="0"/><c:showVal val="0"/>
          <c:showCatName val="0"/><c:showSerName val="0"/>
          <c:showPercent val="0"/><c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:axId val="${ax1}"/><c:axId val="${ax2}"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="${ax1}"/><c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/>
        <c:spPr><a:ln><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></c:spPr>
        <c:txPr><a:bodyPr/><a:lstStyle/>
          <a:p><a:pPr><a:defRPr b="1" sz="900"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:defRPr></a:pPr></a:p>
        </c:txPr>
        <c:crossAx val="${ax2}"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="${ax2}"/>
        <c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="5"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:spPr><a:ln><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="${ax1}"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="none"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

// ── Referência inline DrawingML → chart ───────────────────────────────────
function _wChartRef(rId, cx, cy, docPrId) {
  return `<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
               distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="${cx}" cy="${cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="${docPrId}" name="Chart${docPrId}"/>
      <wp:cNvGraphicFramePr>
        <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
      </wp:cNvGraphicFramePr>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   r:id="${rId}"/>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r>`;
}

// Parágrafo centrado com gráfico
function _wChartPara(rId, cx, cy, id) {
  return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${_wChartRef(rId,cx,cy,id)}</w:p>`;
}

// ─── HELPERS OOXML comuns ──────────────────────────────────────────────────
const _esc = v => String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const _BG = {E:'D4EDDA',MB:'CCE5FF',B:'D4EDDA',R:'FFF3CD',I:'F8D7DA',
             S:'D4EDDA',NS:'F8D7DA',NR:'E2E3E5','':'FFFFFF'};
const _FG = {E:'155724',MB:'1A4A6E',B:'2D4A1E',R:'856404',I:'721C24',
             S:'155724',NS:'721C24',NR:'383D41','':'000000'};

function _wRun(text,opts={}) {
  return `<w:r><w:rPr>${opts.b?'<w:b/>':''}${opts.color?`<w:color w:val="${opts.color}"/>`:''}` +
    `<w:sz w:val="${opts.sz||16}"/></w:rPr><w:t xml:space="preserve">${_esc(text)}</w:t></w:r>`;
}
function _wPara(runs,jc='left',sp={}) {
  return `<w:p><w:pPr><w:jc w:val="${jc}"/>` +
    (sp.before||sp.after?`<w:spacing${sp.before?` w:before="${sp.before}"`:''}${sp.after?` w:after="${sp.after}"` :''}/>`:'')+
    `</w:pPr>${runs}</w:p>`;
}
function _wPB() { return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`; }
function _wCaption(t) { return _wPara(_wRun(t,{sz:16,color:'888888'}),'center',{after:'40'}); }
function _wTc(text,opts={}) {
  const bg=opts.bg||'FFFFFF',fg=opts.fg||'000000',w=opts.w||1200,jc=opts.jc||'left';
  return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>
    <w:shd w:val="clear" w:color="auto" w:fill="${bg}"/></w:tcPr>
    <w:p><w:pPr><w:jc w:val="${jc}"/><w:spacing w:before="0" w:after="0"/></w:pPr>
    <w:r><w:rPr>${opts.b?'<w:b/>':''}<w:color w:val="${fg}"/><w:sz w:val="${opts.sz||16}"/></w:rPr>
    <w:t xml:space="preserve">${_esc(text)}</w:t></w:r></w:p></w:tc>`;
}
function _wTh(text,w=1200) { return _wTc(text,{bg:'2D4A1E',fg:'C9A227',w,jc:'center',b:true,sz:16}); }
function _wTcM(v) { return _wTc(v||'',{bg:_BG[v||'']||'FFFFFF',fg:_FG[v||'']||'000000',w:580,jc:'center',b:true,sz:15}); }

// Tabela de 2 colunas sem bordas (para colocar gráficos lado a lado)
function _tbl2col(c1, c2) {
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders><w:top w:val="none"/><w:bottom w:val="none"/>
      <w:left w:val="none"/><w:right w:val="none"/>
      <w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>
    <w:tr>
      <w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr>${c1}</w:tc>
      <w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr>${c2}</w:tc>
    </w:tr></w:tbl>`;
}

// Tabela de N colunas sem bordas (para OIIs)
function _tblNCols(cells) {
  const pct = Math.floor(100/cells.length);
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders><w:top w:val="none"/><w:bottom w:val="none"/>
      <w:left w:val="none"/><w:right w:val="none"/>
      <w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>
    <w:tr>
      ${cells.map(c=>`<w:tc><w:tcPr><w:tcW w:w="${pct}" w:type="pct"/></w:tcPr>${c}</w:tc>`).join('')}
    </w:tr></w:tbl>`;
}

// Content-Types entry para cada chart
function _chartContentType(n) {
  return `<Override PartName="/word/charts/chart${n}.xml"
    ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
}

// ─── EXPORTAR PDF GERAL ─────────────────────────────────────────────────────
// ════════════════════════════════════════════════════════════════════════
// SVG INLINE — para PDF e impressão (independe de canvas/DOM)
// ════════════════════════════════════════════════════════════════════════

function svgBarras(labels, valores, cores, W, H) {
  const pad = {t:24,r:10,b:28,l:28};
  const gW = W-pad.l-pad.r, gH = H-pad.t-pad.b;
  const maxV = Math.max(...valores,1);
  const barW = Math.floor(gW/labels.length*0.55);
  const gap  = Math.floor(gW/labels.length);
  let rects='',texts='';
  labels.forEach((lbl,i)=>{
    const v=valores[i], bH=Math.round((v/maxV)*gH);
    const x=pad.l+i*gap+(gap-barW)/2, y=pad.t+gH-bH;
    rects+=`<rect x="${x}" y="${y}" width="${barW}" height="${bH}" fill="${cores[i]}" rx="2"/>`;
    if(v>0) texts+=`<text x="${x+barW/2}" y="${y-4}" text-anchor="middle" font-size="11" font-weight="bold" fill="#333">${v}</text>`;
    texts+=`<text x="${x+barW/2}" y="${H-pad.b+13}" text-anchor="middle" font-size="11" fill="#555">${lbl}</text>`;
  });
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">
    <rect width="${W}" height="${H}" fill="white"/>
    <line x1="${pad.l}" y1="${pad.t+gH}" x2="${W-pad.r}" y2="${pad.t+gH}" stroke="#ccc" stroke-width="1"/>
    ${rects}${texts}</svg>`;
}

function svgPizza(labels, valores, cores, total, W, H) {
  if(!total) return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}"><rect width="${W}" height="${H}" fill="white"/><text x="${W/2}" y="${H/2}" text-anchor="middle" fill="#aaa" font-size="12">Sem dados</text></svg>`;
  const cx=W*0.37, cy=H/2, r=Math.min(cx,cy)-8;
  let slices='',ang=-Math.PI/2,leg='';
  valores.forEach((v,i)=>{
    if(!v) return;
    const slice=(v/total)*2*Math.PI;
    const x1=cx+r*Math.cos(ang), y1=cy+r*Math.sin(ang);
    ang+=slice;
    const x2=cx+r*Math.cos(ang), y2=cy+r*Math.sin(ang);
    const large=slice>Math.PI?1:0;
    slices+=`<path d="M${cx},${cy} L${x1.toFixed(1)},${y1.toFixed(1)} A${r},${r},0,${large},1,${x2.toFixed(1)},${y2.toFixed(1)} Z" fill="${cores[i]}" stroke="white" stroke-width="1.5"/>`;
  });
  const ly0=cy-(labels.length*17)/2;
  labels.forEach((lbl,i)=>{
    if(!valores[i]) return;
    const pct=(valores[i]/total*100).toFixed(1), lx=cx+r+12, ly=ly0+i*18;
    leg+=`<rect x="${lx}" y="${ly}" width="11" height="11" fill="${cores[i]}" rx="1"/>`;
    leg+=`<text x="${lx+15}" y="${ly+9}" font-size="10" fill="#333">${lbl}  ${pct}%</text>`;
  });
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">
    <rect width="${W}" height="${H}" fill="white"/>
    ${slices}${leg}</svg>`;
}

function svgLinhaEvolucao(labels, mencoes, W, H) {
  const MVAL={E:5,MB:4,B:3,R:2,I:1}, MCOR={E:'#1565C0',MB:'#2e7d32',B:'#f9a825',R:'#e65100',I:'#c62828'};
  const MLBL=['','I','R','B','MB','E'];
  const pad={t:22,r:16,b:30,l:38}, gW=W-pad.l-pad.r, gH=H-pad.t-pad.b, n=labels.length;
  let grid='',yLbl='';
  for(let v=1;v<=5;v++){
    const y=pad.t+gH-((v-1)/4)*gH;
    grid+=`<line x1="${pad.l}" y1="${y}" x2="${W-pad.r}" y2="${y}" stroke="#eee" stroke-width="1"/>`;
    yLbl+=`<text x="${pad.l-5}" y="${y+4}" text-anchor="end" font-size="10" fill="#888">${MLBL[v]}</text>`;
  }
  const axX=`<line x1="${pad.l}" y1="${pad.t+gH}" x2="${W-pad.r}" y2="${pad.t+gH}" stroke="#bbb" stroke-width="1"/>`;
  const axY=`<line x1="${pad.l}" y1="${pad.t}" x2="${pad.l}" y2="${pad.t+gH}" stroke="#bbb" stroke-width="1"/>`;
  const pts=mencoes.map((m,i)=>({
    x:n>1?pad.l+(i/(n-1))*gW:pad.l+gW/2,
    y:pad.t+gH-((MVAL[m]||0)-1)/4*gH, m
  }));
  let area='',line='',pontos='',xLbl='';
  if(n>1){
    area=`M${pts[0].x},${pad.t+gH} `+pts.map(p=>`L${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ')+` L${pts[n-1].x},${pad.t+gH} Z`;
    line=`M${pts[0].x.toFixed(1)},${pts[0].y.toFixed(1)} `+pts.slice(1).map(p=>`L${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
  }
  pts.forEach((p,i)=>{
    const cor=MCOR[p.m]||'#999';
    pontos+=`<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="5" fill="${cor}" stroke="white" stroke-width="1.5"/>`;
    if(p.m) pontos+=`<text x="${p.x.toFixed(1)}" y="${(p.y-9).toFixed(1)}" text-anchor="middle" font-size="10" font-weight="bold" fill="${cor}">${p.m}</text>`;
    // Quebra label longo em 2 linhas
    const lbl = labels[i] || '';
    const partes = lbl.length > 10 ? [lbl.slice(0,lbl.lastIndexOf(' ')||10), lbl.slice(lbl.lastIndexOf(' ')||10)] : [lbl];
    partes.forEach((parte,pi) => {
      xLbl+=`<text x="${p.x.toFixed(1)}" y="${H-pad.b+13+pi*11}" text-anchor="middle" font-size="9" fill="#555">${parte}</text>`;
    });
  });
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">
    <rect width="${W}" height="${H}" fill="white"/>
    ${grid}${axX}${axY}${yLbl}
    ${area?`<path d="${area}" fill="#2d4a1e22"/>`:''}
    ${line?`<path d="${line}" fill="none" stroke="#2d4a1e" stroke-width="2.5" stroke-linejoin="round"/>`:''}
    ${pontos}${xLbl}</svg>`;
}

// ════════════════════════════════════════════════════════════════════════
// RELATÓRIO GERAL — PDF (A4 landscape, 2 páginas)
// ════════════════════════════════════════════════════════════════════════
function gerarRelatorioPDF(aval, lista) {
  const sufS =lista.filter(r=>r.suficiencia==='S').length;
  const sufNS=lista.filter(r=>r.suficiencia==='NS').length;
  const sufNR=lista.filter(r=>r.suficiencia==='NR'||!r.suficiencia).length;
  const total=lista.length;
  const OM=['E','MB','B','R','I'], CM=['#1565C0','#2e7d32','#f9a825','#e65100','#c62828'], CS=['#2e7d32','#c62828','#9e9e9e'];
  const cM={}; OM.forEach((m,i)=>cM[m]=lista.filter(r=>r.conceito_global===m).length);
  const OIIs=[
    {campo:'conceito_corrida',label:'Corrida 12min'},
    {campo:'conceito_flexao_bracos',label:'Flexão de Braço'},
    {campo:'conceito_abdominal',label:'Abdominal'},
    {campo:'conceito_barra',label:'Barra Fixa'},
  ];
  const cOII={}; OIIs.forEach(o=>{cOII[o.campo]=OM.map(m=>lista.filter(r=>r[o.campo]===m).length);});

  const svgSQ=svgBarras(['S','NS','NR'],[sufS,sufNS,sufNR],CS,320,170);
  const svgSP=svgPizza(['S','NS','NR'],[sufS,sufNS,sufNR],CS,total,260,170);
  const svgMQ=svgBarras(OM,OM.map(m=>cM[m]),CM,320,170);
  const svgMP=svgPizza(OM,OM.map(m=>cM[m]),CM,total,260,170);
  const svgOII=OIIs.map(o=>({
    bar: svgBarras(OM,cOII[o.campo],CM,240,140),
    pie: svgPizza(OM,cOII[o.campo],CM,cOII[o.campo].reduce((s,v)=>s+v,0),200,140),
    label: o.label
  }));

  const win=window.open('','_blank');
  if(!win){toast('Permita popups para gerar PDF.',true);return;}

  const badgeStyle=`.E{background:#d4edda;color:#155724}.MB{background:#cce5ff;color:#1a4a6e}.B{background:#d4edda;color:#2d4a1e}.R{background:#fff3cd;color:#856404}.I{background:#f8d7da;color:#721c24}.S{background:#d4edda;color:#155724}.NS{background:#f8d7da;color:#721c24}`;

  win.document.write(`<!DOCTYPE html><html lang="pt-BR"><head>
  <meta charset="UTF-8"><title>Relatório TAF — ${aval.descricao}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@400;500&display=swap');
    @page{size:A4 landscape;margin:8mm 10mm;}
    *{margin:0;padding:0;box-sizing:border-box;}
    body{font-family:'Barlow',sans-serif;font-size:10px;color:#1a2010;}
    .hdr{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:2px solid #c9a227;padding-bottom:5px;margin-bottom:8px;}
    h1{font-family:'Barlow Condensed',sans-serif;font-size:18px;color:#2d4a1e;letter-spacing:1px;}
    .sub{font-size:9px;color:#888;}
    .stats{display:flex;gap:8px;margin-bottom:10px;}
    .stat{background:#2d4a1e;border-radius:4px;padding:5px 12px;text-align:center;}
    .sv{font-family:'Barlow Condensed',sans-serif;font-size:20px;font-weight:700;color:#c9a227;}
    .sl{font-size:8px;color:rgba(255,255,255,.6);text-transform:uppercase;letter-spacing:1px;}
    .sec{font-family:'Barlow Condensed',sans-serif;font-size:11px;color:#2d4a1e;border-bottom:1px solid #e0d8c8;padding-bottom:2px;margin:8px 0 5px;text-transform:uppercase;letter-spacing:.5px;}
    .g4{display:flex;gap:8px;margin-bottom:10px;}
    .g4>div{flex:1;text-align:center;}
    .g4 label{display:block;font-size:8px;color:#888;margin-bottom:2px;}
    .g4 svg{width:100%;height:auto;}
    .oii-grid{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-bottom:10px;}
    .oii-box{text-align:center;}
    .oii-box label{display:block;font-size:9px;font-weight:bold;color:#2d4a1e;margin-bottom:2px;}
    .oii-box svg{width:100%;height:auto;}
    .pg-break{page-break-before:always;}
    table{width:100%;border-collapse:collapse;font-size:8.5px;}
    th{background:#2d4a1e;color:#c9a227;padding:4px 5px;font-family:'Barlow Condensed',sans-serif;font-size:9px;text-align:left;}
    td{padding:3px 5px;border-bottom:1px solid #ede7d8;}
    tr:nth-child(even) td{background:#f9f6f0;}
    .badge{display:inline-block;padding:1px 4px;border-radius:2px;font-weight:700;font-size:8px;}
    ${badgeStyle}
    .portaria{font-size:7px;color:#aaa;margin-top:8px;border-top:1px solid #e0d8c8;padding-top:4px;text-align:center;}
  </style></head><body>
  <div class="hdr">
    <div><h1>RELATÓRIO DE AVALIAÇÃO FÍSICA — TAF-EB</h1>
      <div class="sub">${aval.descricao} · ${labelTipo(aval.tipo)} · ${aval.ano} · Gerado em ${new Date().toLocaleDateString('pt-BR')}</div>
    </div>
    <div class="sub">Portaria EME/C Ex Nº 850/2022 (EB20-D-03.053)</div>
  </div>
  <div class="stats">
    <div class="stat"><div class="sv">${total}</div><div class="sl">Avaliados</div></div>
    <div class="stat"><div class="sv" style="color:#7ddb8a">${sufS}</div><div class="sl">Suficientes</div></div>
    <div class="stat"><div class="sv" style="color:#e87070">${sufNS}</div><div class="sl">NS</div></div>
    <div class="stat"><div class="sv" style="color:#aaa">${sufNR}</div><div class="sl">NR</div></div>
    <div class="stat"><div class="sv" style="color:#c9a227">${total?(sufS/total*100).toFixed(1):0}%</div><div class="sl">% Suficientes</div></div>
  </div>

  <div class="sec">Desempenho Geral — Suficiência e Menção Global</div>
  <div class="g4">
    <div><label>Quantidade / Suficiência</label>${svgSQ}</div>
    <div><label>Percentual / Suficiência</label>${svgSP}</div>
    <div><label>Quantidade / Menção Global</label>${svgMQ}</div>
    <div><label>Percentual / Menção Global</label>${svgMP}</div>
  </div>

  <div class="sec">Desempenho por OII — Menção (barras + percentual)</div>
  <div class="oii-grid">
    ${svgOII.map(o=>`<div class="oii-box">
      <label>${o.label}</label>
      ${o.bar}
      <div style="margin-top:4px;border-top:1px solid #f0ede6;padding-top:4px">${o.pie}</div>
    </div>`).join('')}
  </div>

  <div class="pg-break"></div>
  <div class="hdr"><div><h1>LISTA INDIVIDUAL — ${aval.descricao}</h1></div></div>
  <table><thead><tr>
    <th>#</th><th>Nome de Guerra</th><th>Posto</th><th>Sx</th>
    <th>Corrida</th><th>M</th><th>Flexão</th><th>M</th>
    <th>Abdominal</th><th>M</th><th>Barra</th><th>M</th>
    <th>Global</th><th>Sufic.</th><th>Cham.</th>
  </tr></thead><tbody>
    ${lista.map((r,i)=>`<tr>
      <td>${i+1}</td>
      <td><strong>${r.nome_guerra||r.nome}</strong></td>
      <td>${r.posto_graduacao}</td><td>${r.sexo}</td>
      <td>${r.corrida_12min||'—'}</td><td><span class="badge ${r.conceito_corrida||''}">${r.conceito_corrida||''}</span></td>
      <td>${r.flexao_bracos||'—'}</td><td><span class="badge ${r.conceito_flexao_bracos||''}">${r.conceito_flexao_bracos||''}</span></td>
      <td>${r.abdominal_supra||'—'}</td><td><span class="badge ${r.conceito_abdominal||''}">${r.conceito_abdominal||''}</span></td>
      <td>${r.flexao_barra||'—'}</td><td><span class="badge ${r.conceito_barra||''}">${r.conceito_barra||''}</span></td>
      <td><span class="badge ${r.conceito_global||''}">${r.conceito_global||'NR'}</span></td>
      <td><span class="badge ${r.suficiencia||''}">${r.suficiencia||'NR'}</span></td>
      <td>${r.chamada}ª</td>
    </tr>`).join('')}
  </tbody></table>
  <div class="portaria">Portaria EME/C Ex Nº 850, de 31 de agosto de 2022 (EB20-D-03.053) · TAF-EB Sistema de Avaliação Física</div>
  <script>window.onload=function(){setTimeout(function(){window.print();},400);};</script>
  </body></html>`);
  win.document.close();
}


async function exportarPDF(avalId) {
  const aval  = state.avaliacoes.find(a => a.id === avalId);
  const lista = await window.api.resultados.listar(avalId);
  if (!lista.length) { toast('Nenhum resultado para exportar.', true); return; }
  toast('Abrindo PDF...');
  gerarRelatorioPDF(aval, lista);
}

// ─── FICHAS INDIVIDUAIS EM LOTE — PDF ──────────────────────────────────────
async function exportarFichasLotePDF(avalId) {
  const aval  = state.avaliacoes.find(a => a.id === avalId);
  const lista = await window.api.resultados.listar(avalId);
  if (!lista.length) { toast('Nenhum resultado para gerar fichas.', true); return; }
  toast(`Gerando ${lista.length} fichas PDF...`);

  const ORDEM_TIPO = ['diagnostica','1taf','2taf','3taf'];
  const militaresIds = [...new Set(lista.map(r => r.militar_id))];
  const dadosMilitares = await Promise.all(
    militaresIds.map(async mid => {
      const m   = state.militares.find(x => x.id === mid);
      const res = await window.api.resultados.porMilitar(mid);
      return { m, res };
    })
  );
  dadosMilitares.sort((a,b) => (a.m.nome_guerra||a.m.nome).localeCompare(b.m.nome_guerra||b.m.nome));

  const fichasHTML = dadosMilitares.map(({ m, res }) => {
    res.sort((a,b) => a.ano - b.ano || ORDEM_TIPO.indexOf(a.tipo) - ORDEM_TIPO.indexOf(b.tipo));

    const labels = res.map(r => r.tipo === 'diagnostica'
      ? `AD ${r.ano}`
      : (r.descricao || r.tipo.replace('taf','ºTAF')+' '+r.ano));
    const armaObj = getArma(m.arma);
    const padrao  = calcularPadrao(m.arma, m.situacao_funcional);
    const fotoSrc = _fotosCache[m.id] || null;

    const OIIs = [
      { campo:'conceito_barra',         label:'Barra Fixa' },
      { campo:'conceito_flexao_bracos', label:'Flexão de Braço' },
      { campo:'conceito_abdominal',     label:'Abdominal' },
      { campo:'conceito_corrida',       label:'Corrida 12min' },
    ];

    const svgGlobal = res.length > 0
      ? svgLinhaEvolucao(labels, res.map(r => r.conceito_global), 740, 140)
      : '<p style="color:#aaa;text-align:center;padding:10px">Sem dados</p>';

    const svgsOII = OIIs.map(o => svgLinhaEvolucao(labels, res.map(r => r[o.campo]), 230, 120));

    const tabelaLinhas = res.map(r => `
      <tr>
        <td>${r.descricao}</td><td>${labelTipo(r.tipo)}</td><td>${r.ano}</td><td>${r.chamada}ª</td>
        <td>${r.corrida_12min||'—'}</td><td class="m">${r.conceito_corrida||''}</td>
        <td>${r.flexao_bracos||'—'}</td><td class="m">${r.conceito_flexao_bracos||''}</td>
        <td>${r.abdominal_supra||'—'}</td><td class="m">${r.conceito_abdominal||''}</td>
        <td>${r.flexao_barra||'—'}</td><td class="m">${r.conceito_barra||''}</td>
        <td class="g">${r.conceito_global||'NR'}</td><td class="s">${r.suficiencia||'NR'}</td>
      </tr>`).join('');

    return `
    <div class="ficha">
      <div class="topo">
        ${fotoSrc ? `<img class="foto" src="${fotoSrc}">` : '<div class="foto-ph">👤</div>'}
        <div class="dados">
          <div class="nome">${m.posto_graduacao} ${armaObj?.sigla||''} ${m.nome_guerra||m.nome}</div>
          <div class="sub">${m.nome}</div>
          <div class="dgrid">
            <span><b>Nasc.:</b> ${formatarData(m.data_nascimento)}</span>
            <span><b>Idade:</b> ${TAF.calcularIdade(m.data_nascimento)} anos</span>
            <span><b>Sexo:</b> ${m.sexo==='M'?'Masculino':'Feminino'}</span>
            <span><b>Arma:</b> ${armaObj?.sigla||m.arma||'—'} — ${armaObj?.label||''}</span>
            <span><b>LEM:</b> ${m.lem||'—'}</span>
            <span><b>Padrão:</b> ${padrao}</span>
            <span><b>OM:</b> ${m.om||'—'}</span>
          </div>
        </div>
        <div class="aval-info">
          <div style="font-size:9px;color:#888">Avaliação</div>
          <div style="font-size:11px;font-weight:700;color:#2d4a1e">${aval.descricao}</div>
          <div style="font-size:9px;color:#888">${aval.ano}</div>
        </div>
      </div>
      <div class="sec">Evolução — Menção Global (${labels.join(' → ')})</div>
      <div class="graf-global">${svgGlobal}</div>
      <div class="sec">Evolução por OII</div>
      <div class="oii-grid">
        ${OIIs.map((o,i)=>`<div class="oii-item"><div class="oii-lbl">${o.label}</div>${svgsOII[i]}</div>`).join('')}
      </div>
      <div class="sec">Histórico</div>
      <table>
        <thead><tr>
          <th>Avaliação</th><th>Tipo</th><th>Ano</th><th>Ch</th>
          <th>Corrida</th><th>M</th><th>Flexão</th><th>M</th>
          <th>Abdom</th><th>M</th><th>Barra</th><th>M</th>
          <th>Global</th><th>Suf</th>
        </tr></thead>
        <tbody>${tabelaLinhas}</tbody>
      </table>
    </div>`;
  }).join('');

  const win = window.open('', '_blank');
  if (!win) { toast('Permita popups para imprimir.', true); return; }
  win.document.write(`<!DOCTYPE html><html lang="pt-BR"><head>
  <meta charset="UTF-8"><title>Fichas — ${aval.descricao}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@400;500&display=swap');
    @page { size: A4 landscape; margin: 7mm 9mm; }
    *{margin:0;padding:0;box-sizing:border-box;}
    body{font-family:'Barlow',sans-serif;font-size:9px;color:#1a2010;}
    .ficha{page-break-after:always;page-break-inside:avoid;}
    .ficha:last-child{page-break-after:auto;}
    .topo{display:flex;align-items:flex-start;gap:10px;border-bottom:2px solid #c9a227;padding-bottom:7px;margin-bottom:7px;}
    .foto{width:60px;height:76px;object-fit:cover;border:1px solid #bbb;border-radius:2px;flex-shrink:0;}
    .foto-ph{width:60px;height:76px;background:#f0ede6;border:1px dashed #ccc;border-radius:2px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:22px;color:#ccc;}
    .dados{flex:1;}.nome{font-family:'Barlow Condensed',sans-serif;font-size:15px;color:#2d4a1e;font-weight:700;}
    .sub{font-size:9px;color:#888;margin-bottom:4px;}.dgrid{display:flex;flex-wrap:wrap;gap:2px 14px;}
    .aval-info{text-align:right;flex-shrink:0;}
    .sec{font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:.5px;color:#2d4a1e;
         border-bottom:1px solid #e0d8c8;padding-bottom:2px;margin:6px 0 4px;text-transform:uppercase;}
    .graf-global svg{width:100%;height:auto;display:block;}
    .oii-grid{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:6px;margin-bottom:6px;}
    .oii-item{text-align:center;}.oii-lbl{font-size:8px;color:#888;margin-bottom:1px;}
    .oii-item svg{width:100%;height:auto;display:block;}
    table{width:100%;border-collapse:collapse;font-size:8px;}
    th{background:#2d4a1e;color:#c9a227;padding:2px 4px;font-family:'Barlow Condensed',sans-serif;font-size:8px;}
    td{padding:2px 4px;border-bottom:1px solid #ede7d8;}
    tr:nth-child(even) td{background:#f9f6f0;}
    td.m{font-weight:700;text-align:center;}td.g{font-weight:700;text-align:center;color:#2d4a1e;}td.s{text-align:center;}
  </style></head><body>
  ${fichasHTML}
  <script>window.onload=function(){setTimeout(function(){window.print();},500);};</script>
  </body></html>`);
  win.document.close();
}


// ─── FUNÇÕES MK — XML de chart DrawingML (axIds únicos, validado) ───────────
function _mkChartBarras(cid, titulo, cats, vals, cores) {
  const ax1=cid*10+1, ax2=cid*10+2, n=cats.length;
  const esc=v=>String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const ptC=cats.map((v,i)=>`<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('');
  const ptV=vals.map((v,i)=>`<c:pt idx="${i}"><c:v>${v||0}</c:v></c:pt>`).join('');
  const dp =cats.map((_,i)=>{const cor=(cores&&cores[i]||'#607D8B').replace('#','');return`<c:dPt><c:idx val="${i}"/><c:spPr><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill></a:ln></c:spPr></c:dPt>`;}).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:lang val="pt-BR"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1" sz="1000"><a:solidFill><a:srgbClr val="2D4A1E"/></a:solidFill></a:rPr><a:t>${titulo}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>${dp}<c:dLbls><c:spPr><a:noFill/></c:spPr><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:cat><c:strRef><c:f/><c:strCache><c:ptCount val="${n}"/>${ptC}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f/><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${ptV}</c:numCache></c:numRef></c:val></c:ser><c:axId val="${ax1}"/><c:axId val="${ax2}"/></c:barChart><c:catAx><c:axId val="${ax1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="${ax2}"/></c:catAx><c:valAx><c:axId val="${ax2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="General" sourceLinked="0"/><c:crossAx val="${ax1}"/></c:valAx></c:plotArea><c:legend><c:legendPos val="none"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function _mkChartPizza(cid, titulo, cats, vals, cores) {
  const n=cats.length;
  const esc=v=>String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const ptC=cats.map((v,i)=>`<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('');
  const ptV=vals.map((v,i)=>`<c:pt idx="${i}"><c:v>${v||0}</c:v></c:pt>`).join('');
  const dp =cats.map((_,i)=>{const cor=(cores&&cores[i]||'#607D8B').replace('#','');return`<c:dPt><c:idx val="${i}"/><c:spPr><a:solidFill><a:srgbClr val="${cor}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:w val="12700"/></a:ln></c:spPr></c:dPt>`;}).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:lang val="pt-BR"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1" sz="1000"><a:solidFill><a:srgbClr val="2D4A1E"/></a:solidFill></a:rPr><a:t>${titulo}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${dp}<c:dLbls><c:spPr><a:noFill/></c:spPr><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:separator>: </c:separator></c:dLbls><c:cat><c:strRef><c:f/><c:strCache><c:ptCount val="${n}"/>${ptC}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f/><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${n}"/>${ptV}</c:numCache></c:numRef></c:val></c:ser><c:firstSliceAng val="270"/></c:pieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function _mkChartLinha(cid, titulo, labels, mencoes) {
  const MVAL={E:5,MB:4,B:3,R:2,I:1};
  const MCOR={E:'1565C0',MB:'2E7D32',B:'558B2F',R:'F9A825',I:'C62828'};
  const ax1=cid*10+1, ax2=cid*10+2, n=labels.length;
  const vals=mencoes.map(m=>MVAL[m]||0);
  const esc=v=>String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const norm=l=>l.replace(/º/g,'o').replace(/ã/g,'a').replace(/ç/g,'c').replace(/õ/g,'o')
              .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i').replace(/ó/g,'o')
              .replace(/ú/g,'u').replace(/â/g,'a').replace(/ê/g,'e').replace(/ô/g,'o');
  const ptC=labels.map((l,i)=>'<c:pt idx="'+i+'"><c:v>'+esc(norm(l))+'</c:v></c:pt>').join('');
  const ptV=vals.map((v,i)=>'<c:pt idx="'+i+'"><c:v>'+v+'</c:v></c:pt>').join('');
  // Pontos coloridos por menção
  const dpts=mencoes.map((m,i)=>{
    if(!m||!MVAL[m])return'';
    const cor=MCOR[m]||'607D8B';
    return '<c:dPt><c:idx val="'+i+'"/>'+
      '<c:spPr><a:solidFill><a:srgbClr val="'+cor+'"/></a:solidFill>'+
      '<a:ln><a:solidFill><a:srgbClr val="'+cor+'"/></a:solidFill></a:ln></c:spPr>'+
      '</c:dPt>';
  }).join('');
  // Rótulo de menção (E, MB, B, R, I) acima de cada ponto
  const dlbls=mencoes.map((m,i)=>{
    if(!m||!MVAL[m])return'';
    const cor=MCOR[m]||'607D8B';
    return '<c:dLbl><c:idx val="'+i+'"/>'+
      '<c:spPr><a:noFill/></c:spPr>'+
      '<c:txPr><a:bodyPr/><a:lstStyle/>'+
        '<a:p><a:pPr><a:defRPr b="1" sz="1000">'+
          '<a:solidFill><a:srgbClr val="'+cor+'"/></a:solidFill>'+
        '</a:defRPr></a:pPr></a:p></c:txPr>'+
      '<c:showLegendKey val="0"/>'+
      '<c:showVal val="0"/>'+
      '<c:showCatName val="0"/>'+
      '<c:showSerName val="0"/>'+
      '<c:showPercent val="0"/>'+
      '<c:showBubbleSize val="0"/>'+
      '<c:tx><c:strRef><c:f/><c:strCache><c:ptCount val="1"/>'+
        '<c:pt idx="0"><c:v>'+m+'</c:v></c:pt>'+
      '</c:strCache></c:strRef></c:tx>'+
      '</c:dLbl>';
  }).join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'+
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'+
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+
    '<c:lang val="pt-BR"/>'+
    '<c:chart><c:autoTitleDeleted val="1"/>'+
    '<c:plotArea><c:layout/>'+
    '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'+
    '<c:ser><c:idx val="0"/><c:order val="0"/>'+
    '<c:spPr><a:noFill/>'+
    '<a:ln w="25400"><a:solidFill><a:srgbClr val="2D4A1E"/></a:solidFill></a:ln>'+
    '</c:spPr>'+
    '<c:marker><c:symbol val="circle"/><c:size val="7"/>'+
    '<c:spPr><a:solidFill><a:srgbClr val="2D4A1E"/></a:solidFill></c:spPr>'+
    '</c:marker>'+
    dpts+
    '<c:dLbls>'+dlbls+
    '<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>'+
    '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>'+
    '</c:dLbls>'+
    '<c:cat><c:strRef><c:f/><c:strCache><c:ptCount val="'+n+'"/>'+ptC+'</c:strCache></c:strRef></c:cat>'+
    '<c:val><c:numRef><c:f/><c:numCache><c:formatCode>General</c:formatCode>'+
    '<c:ptCount val="'+n+'"/>'+ptV+'</c:numCache></c:numRef></c:val>'+
    '<c:smooth val="0"/></c:ser>'+
    '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>'+
    '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'+
    '<c:axId val="'+ax1+'"/><c:axId val="'+ax2+'"/></c:lineChart>'+
    '<c:catAx><c:axId val="'+ax1+'"/><c:scaling><c:orientation val="minMax"/></c:scaling>'+
    '<c:delete val="0"/><c:axPos val="b"/>'+
    '<c:spPr><a:ln><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></c:spPr>'+
    '<c:crossAx val="'+ax2+'"/></c:catAx>'+
    '<c:valAx><c:axId val="'+ax2+'"/>'+
    '<c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="5"/></c:scaling>'+
    '<c:delete val="1"/><c:axPos val="l"/>'+
    '<c:numFmt formatCode="General" sourceLinked="0"/>'+
    '<c:spPr><a:ln><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></c:spPr>'+
    '<c:crossAx val="'+ax1+'"/></c:valAx>'+
    '</c:plotArea>'+
    '<c:legend><c:legendPos val="none"/></c:legend>'+
    '<c:plotVisOnly val="1"/>'+
    '</c:chart></c:chartSpace>';
}



// ─── HELPERS OOXML ──────────────────────────────────────────────────────────
function _oe(v){return String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
const _OBG={E:'D4EDDA',MB:'CCE5FF',B:'D4EDDA',R:'FFF3CD',I:'F8D7DA',S:'D4EDDA',NS:'F8D7DA',NR:'E2E3E5','':'FFFFFF'};
const _OFG={E:'155724',MB:'1A4A6E',B:'2D4A1E',R:'856404',I:'721C24',S:'155724',NS:'721C24',NR:'383D41','':'000000'};
function _oRun(t,b,color,sz){return`<w:r><w:rPr>${b?'<w:b/>':''}${color?`<w:color w:val="${color}"/>`:''}${sz?`<w:sz w:val="${sz}"/>`:'<w:sz w:val="20"/>'}</w:rPr><w:t xml:space="preserve">${_oe(t)}</w:t></w:r>`;}
function _oP(runs,jc,bef,aft){return`<w:p><w:pPr><w:jc w:val="${jc||'left'}"/>${(bef||aft)?`<w:spacing${bef?` w:before="${bef}"`:''}${aft?` w:after="${aft}"`:''}/>`:''}` + `</w:pPr>${runs}</w:p>`;}
function _oPB(){return`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;}
function _oTc(t,w,jc,bg,fg,b,sz){return`<w:tc><w:tcPr><w:tcW w:w="${w||1200}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="${bg||'FFFFFF'}"/></w:tcPr><w:p><w:pPr><w:jc w:val="${jc||'left'}"/><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr>${b?'<w:b/>':''}<w:color w:val="${fg||'000000'}"/><w:sz w:val="${sz||16}"/></w:rPr><w:t xml:space="preserve">${_oe(t)}</w:t></w:r></w:p></w:tc>`;}
function _oTh(t,w){return _oTc(t,w||1200,'center','2D4A1E','C9A227',true,16);}
function _oTcM(v){return _oTc(v||'',580,'center',_OBG[v||'']||'FFFFFF',_OFG[v||'']||'000000',true,15);}
function _oChart(rId,cx,cy,pid){return`<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${pid}" name="Chart${pid}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rId}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;}
function _oChartP(rId,cx,cy,pid){return`<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${_oChart(rId,cx,cy,pid)}</w:p>`;}
function _oTbl2(a,b){return`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr><w:tr><w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr>${a}</w:tc><w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr>${b}</w:tc></w:tr></w:tbl>`;}
function _oTbl4(cells){return`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="none"/><w:bottom w:val="none"/><w:left w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr><w:tr>${cells.map(c=>`<w:tc><w:tcPr><w:tcW w:w="25" w:type="pct"/></w:tcPr>${c}</w:tc>`).join('')}</w:tr></w:tbl>`;}
function _oDocSkeleton(body){return`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body>${body}</w:body></w:document>`;}
function _oDocSkeletonLandscape(body,m){const mg=m||720;return`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><w:body>${body}<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="${mg}" w:right="${mg}" w:bottom="${mg}" w:left="${mg}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`;}
function _oZipMeta(charts,extra){
  const rels_c=charts.map(ch=>`<Relationship Id="rChart${ch.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart${ch.id}.xml"/>`).join('\n  ');
  const ct_c  =charts.map(ch=>`<Override PartName="/word/charts/chart${ch.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join('\n  ');
  const docRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n  ${rels_c}\n</Relationships>`;
  const ct=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n  <Default Extension="xml" ContentType="application/xml"/>\n  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n  ${ct_c}\n</Types>`;
  const rr=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n</Relationships>`;
  const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>\n  <w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>\n</w:styles>`;
  const settings=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>\n</w:settings>`;
  return {docRels,ct,rr,styles,settings};
}
async function _oSalvarDocx(zip, filename) {
  const blob = await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href=url; a.download=filename; a.click();
  URL.revokeObjectURL(url);
}

// ─── RELATÓRIO GERAL — Word com gráficos nativos ────────────────────────────
async function exportarWord(avalId) {
  if (typeof JSZip==='undefined'){toast('JSZip não carregado.',true);return;}
  const aval = state.avaliacoes.find(a=>a.id===avalId);
  const lista= await window.api.resultados.listar(avalId);
  if (!lista.length){toast('Nenhum resultado para exportar.',true);return;}
  toast('Gerando Word com gráficos nativos...');

  const sufS=lista.filter(r=>r.suficiencia==='S').length;
  const sufNS=lista.filter(r=>r.suficiencia==='NS').length;
  const sufNR=lista.filter(r=>!r.suficiencia||r.suficiencia==='NR').length;
  const total=lista.length;
  const OM=['E','MB','B','R','I'];
  const CM=[_CM.E,_CM.MB,_CM.B,_CM.R,_CM.I];
  const CS=[_CM.S,_CM.NS,_CM.NR];
  const cMen={}; OM.forEach(m=>cMen[m]=lista.filter(r=>r.conceito_global===m).length);
  const OIIs=[
    {campo:'conceito_corrida',label:'Corrida 12min'},
    {campo:'conceito_flexao_bracos',label:'Flexão de Braço'},
    {campo:'conceito_abdominal',label:'Abdominal'},
    {campo:'conceito_barra',label:'Barra Fixa'},
  ];
  const cOII={}; OIIs.forEach(o=>{cOII[o.campo]=OM.map(m=>lista.filter(r=>r[o.campo]===m).length);});

  const charts=[
    {id:1,xml:_mkChartBarras(1,'Qtd / Suficiência',['S','NS','NR'],[sufS,sufNS,sufNR],CS)},
    {id:2,xml:_mkChartPizza (2,'Pct / Suficiência',['S','NS','NR'],[sufS,sufNS,sufNR],CS)},
    {id:3,xml:_mkChartBarras(3,'Qtd / Menção Global',OM,OM.map(m=>cMen[m]),CM)},
    {id:4,xml:_mkChartPizza (4,'Pct / Menção Global',OM,OM.map(m=>cMen[m]),CM)},
    ...OIIs.map((o,i)=>({id:5+i,xml:_mkChartBarras(5+i,o.label,OM,cOII[o.campo],CM)})),
  ];

  // Linhas da tabela
  const linhas=lista.map((r,i)=>`<w:tr>${_oTc(i+1,380,'center','FFFFFF','000000',false,15)}${_oTc(r.nome_guerra||r.nome,1800,'left','FFFFFF','000000',true,15)}${_oTc(r.posto_graduacao,860,'left','FFFFFF','000000',false,15)}${_oTc(r.sexo,340,'center','FFFFFF','000000',false,15)}${_oTc(r.corrida_12min||'',650,'center','FFFFFF','000000',false,15)}${_oTcM(r.conceito_corrida)}${_oTc(r.flexao_bracos||'',650,'center','FFFFFF','000000',false,15)}${_oTcM(r.conceito_flexao_bracos)}${_oTc(r.abdominal_supra||'',650,'center','FFFFFF','000000',false,15)}${_oTcM(r.conceito_abdominal)}${_oTc(r.flexao_barra||'',650,'center','FFFFFF','000000',false,15)}${_oTcM(r.conceito_barra)}${_oTcM(r.conceito_global)}${_oTcM(r.suficiencia)}${_oTc(r.chamada+'ª',460,'center','FFFFFF','000000',false,15)}</w:tr>`).join('');

  const body=
    _oP(_oRun('RELATÓRIO DE AVALIAÇÃO FÍSICA — TAF-EB',true,'2D4A1E',36),'left',0,50)+
    _oP(_oRun(`${aval.descricao} · ${labelTipo(aval.tipo)} · ${aval.ano} · ${new Date().toLocaleDateString('pt-BR')}`,false,'888888',18),'left',0,40)+
    _oP(_oRun(`Total: ${total}   Suficientes: ${sufS} (${total?(sufS/total*100).toFixed(1):0}%)   NS: ${sufNS}   NR: ${sufNR}`,true,'',18),'left',0,80)+
    _oP(_oRun('SUFICIÊNCIA',true,'2D4A1E',22),'left',0,40)+
    _oTbl2(
      _oP(_oRun('Qtd',false,'888888',14),'center',0,10)+_oChartP('rChart1',3200000,2000000,1),
      _oP(_oRun('Pct',false,'888888',14),'center',0,10)+_oChartP('rChart2',3200000,2000000,2)
    )+
    _oTbl2(
      _oP(_oRun('Qtd / Menção Global',false,'888888',14),'center',0,10)+_oChartP('rChart3',3200000,2000000,3),
      _oP(_oRun('Pct / Menção Global',false,'888888',14),'center',0,10)+_oChartP('rChart4',3200000,2000000,4)
    )+
    _oP(_oRun('OII',true,'2D4A1E',22),'left',120,40)+
    _oTbl4(OIIs.map((o,i)=>_oP(_oRun(o.label,false,'888888',14),'center',0,10)+_oChartP(`rChart${5+i}`,2200000,1600000,5+i)))+
    _oPB()+
    _oP(_oRun(`LISTA — ${aval.descricao}`,true,'2D4A1E',28),'left',0,60)+
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="EDE7D8"/><w:insideV w:val="single" w:sz="4" w:color="EDE7D8"/></w:tblBorders></w:tblPr><w:tr>${_oTh('#',380)}${_oTh('Nome de Guerra',1800)}${_oTh('Posto',860)}${_oTh('Sx',340)}${_oTh('Corrida',650)}${_oTh('M',580)}${_oTh('Flexão',650)}${_oTh('M',580)}${_oTh('Abdom',650)}${_oTh('M',580)}${_oTh('Barra',650)}${_oTh('M',580)}${_oTh('Global',580)}${_oTh('Suf.',580)}${_oTh('Ch.',460)}</w:tr>${linhas}</w:tbl>`;

  const docXml=_oDocSkeletonLandscape(body,720);
  const meta=_oZipMeta(charts);
  const zip=new JSZip();
  zip.file('[Content_Types].xml',meta.ct);
  zip.file('_rels/.rels',meta.rr);
  zip.file('word/document.xml',docXml);
  zip.file('word/_rels/document.xml.rels',meta.docRels);
  zip.file('word/styles.xml',meta.styles);
  zip.file('word/settings.xml',meta.settings);
  charts.forEach(ch=>zip.file(`word/charts/chart${ch.id}.xml`,ch.xml));
  await _oSalvarDocx(zip,`relatorio-${aval.descricao.replace(/[^a-zA-Z0-9]/g,'-')}.docx`);
  toast('✔ Word gerado!');
}

// ─── FICHAS INDIVIDUAIS EM LOTE — Word ──────────────────────────────────────


async function exportarFichasLoteWord(avalId) {
  if (typeof JSZip==='undefined'){toast('JSZip nao carregado.',true);return;}
  const aval=state.avaliacoes.find(a=>a.id===avalId);
  const lista=await window.api.resultados.listar(avalId);
  if(!lista.length){toast('Nenhum resultado.',true);return;}

  const ORDEM=['diagnostica','1taf','2taf','3taf'];
  const mids=[...new Set(lista.map(r=>r.militar_id))];
  const dados=await Promise.all(mids.map(async mid=>{
    const m=state.militares.find(x=>x.id===mid);
    const res=await window.api.resultados.porMilitar(mid);
    return {m,res};
  }));
  dados.sort((a,b)=>(a.m.nome_guerra||a.m.nome).localeCompare(b.m.nome_guerra||b.m.nome));
  toast('Gerando fichas Word ('+dados.length+')...');

  // ── Configurações de cor e valor por menção ─────────────────────────────
  const MVAL={E:5,MB:4,B:3,R:2,I:1,NR:0};
  const MCOR={E:'#1565C0',MB:'#2E7D32',B:'#558B2F',R:'#F9A825',I:'#C62828',NR:'#9E9E9E'};
  const YLABELS=['NR','I','R','B','MB','E'];
  const OBG={E:'D4EDDA',MB:'CCE5FF',B:'D4EDDA',R:'FFF3CD',I:'F8D7DA',S:'D4EDDA',NS:'F8D7DA',NR:'E2E3E5','':'FFFFFF'};
  const OFG={E:'155724',MB:'1A4A6E',B:'2D4A1E',R:'856404',I:'721C24',S:'155724',NS:'721C24',NR:'383D41','':'000000'};
  const OIIs=[
    {campo:'conceito_barra',        label:'Barra Fixa'},
    {campo:'conceito_flexao_bracos',label:'Flexao de Braco'},
    {campo:'conceito_abdominal',    label:'Abdominal'},
    {campo:'conceito_corrida',      label:'Corrida 12min'},
  ];

  // ── Gerar gráfico PNG via Canvas ────────────────────────────────────────
  function gerarPNG(labels, mencoes, titulo, W, H) {
    const canvas = document.createElement('canvas');
    const DPR = 2; // alta resolução
    canvas.width  = W * DPR;
    canvas.height = H * DPR;
    const ctx = canvas.getContext('2d');
    ctx.scale(DPR, DPR);

    const PAD = {top:28, right:16, bottom:36, left:40};
    const pw = W - PAD.left - PAD.right;
    const ph = H - PAD.top  - PAD.bottom;

    // Fundo branco
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);

    // Título
    if (titulo) {
      ctx.fillStyle = '#888888';
      ctx.font = 'bold 10px Calibri, Arial';
      ctx.textAlign = 'center';
      ctx.fillText(titulo, W/2, 14);
    }

    // Grid horizontal + labels eixo Y (NR→E)
    ctx.strokeStyle = '#CCCCCC';
    ctx.lineWidth = 0.8;
    for (let i=0;i<=5;i++) {
      const y = PAD.top + ph - (i/5)*ph;
      ctx.beginPath(); ctx.moveTo(PAD.left,y); ctx.lineTo(PAD.left+pw,y); ctx.stroke();
      ctx.fillStyle = '#333333';
      ctx.font = 'bold 10px Calibri, Arial';
      ctx.textAlign = 'right';
      ctx.fillText(YLABELS[i], PAD.left-4, y+4);
    }

    // Linha de base (eixo X)
    ctx.strokeStyle = '#CCCCCC';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(PAD.left, PAD.top+ph);
    ctx.lineTo(PAD.left+pw, PAD.top+ph);
    ctx.stroke();

    const n = labels.length;
    const xPos = i => PAD.left + (n===1 ? pw/2 : (i/(n-1))*pw);
    const yPos = v => PAD.top + ph - (v/5)*ph;
    const vals = mencoes.map(m=>MVAL[m]||0);

    // Linha conectando pontos
    if (n > 1) {
      ctx.strokeStyle = '#2D4A1E';
      ctx.lineWidth = 2.5;
      ctx.beginPath();
      ctx.moveTo(xPos(0), yPos(vals[0]));
      for (let i=1;i<n;i++) ctx.lineTo(xPos(i), yPos(vals[i]));
      ctx.stroke();
    }

    // Pontos + conceito acima
    mencoes.forEach((m,i) => {
      const cor = MCOR[m] || '#607D8B';
      const x = xPos(i), y = yPos(vals[i]);
      // Bolinha
      ctx.beginPath();
      ctx.arc(x, y, 7, 0, Math.PI*2);
      ctx.fillStyle = cor;
      ctx.fill();
      // Conceito acima
      ctx.fillStyle = cor;
      ctx.font = 'bold 11px Calibri, Arial';
      ctx.textAlign = 'center';
      ctx.fillText(m, x, y-11);
    });

    // Labels eixo X
    labels.forEach((lbl,i) => {
      ctx.fillStyle = '#222222';
      ctx.font = 'bold 10px Calibri, Arial';
      ctx.textAlign = 'center';
      ctx.fillText(lbl, xPos(i), PAD.top+ph+16);
    });

    // Converter para PNG base64
    return canvas.toDataURL('image/png').split(',')[1];
  }

  // ── Helpers Word XML ─────────────────────────────────────────────────────
  function esc(v){return String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

  function wTc(t,w,jc,bg,fg,b,sz){
    return '<w:tc><w:tcPr><w:tcW w:w="'+w+'" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="'+bg+'"/></w:tcPr>'+
           '<w:p><w:pPr><w:jc w:val="'+jc+'"/><w:spacing w:before="0" w:after="0"/></w:pPr>'+
           '<w:r><w:rPr>'+(b?'<w:b/>':'')+'<w:color w:val="'+fg+'"/><w:sz w:val="'+sz+'"/></w:rPr>'+
           '<w:t xml:space="preserve">'+esc(t)+'</w:t></w:r></w:p></w:tc>';
  }
  function wTh(t,w){return wTc(t,w,'center','2D4A1E','C9A227',true,16);}
  function wTcM(v){return wTc(v||'',500,'center',OBG[v||'']||'FFFFFF',OFG[v||'']||'000000',true,15);}
  function wP(t,jc,color,sz,b,bef,aft){
    return '<w:p><w:pPr><w:jc w:val="'+jc+'"/><w:spacing w:before="'+(bef||0)+'" w:after="'+(aft||60)+'"/></w:pPr>'+
           '<w:r><w:rPr>'+(b?'<w:b/>':'')+'<w:color w:val="'+color+'"/><w:sz w:val="'+sz+'"/></w:rPr>'+
           '<w:t xml:space="preserve">'+esc(t)+'</w:t></w:r></w:p>';
  }
  function wImg(rid, cx, cy, pid){
    return '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'+
      '<wp:inline distT="0" distB="0" distL="0" distR="0">'+
      '<wp:extent cx="'+cx+'" cy="'+cy+'"/>'+
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>'+
      '<wp:docPr id="'+pid+'" name="Img'+pid+'"/>'+
      '<wp:cNvGraphicFramePr/>'+
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'+
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'+
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'+
      '<pic:nvPicPr><pic:cNvPr id="'+pid+'" name="Img'+pid+'"/><pic:cNvPicPr/></pic:nvPicPr>'+
      '<pic:blipFill>'+
      '<a:blip r:embed="'+rid+'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'+
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'+
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm>'+
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'+
      '</pic:pic></a:graphicData></a:graphic>'+
      '</wp:inline></w:drawing></w:r>';
  }
  function wTbl2(a,b){
    return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'+
      '<w:tblBorders><w:top w:val="none"/><w:bottom w:val="none"/>'+
      '<w:left w:val="none"/><w:right w:val="none"/>'+
      '<w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>'+
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr>'+a+'</w:tc>'+
      '<w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr>'+b+'</w:tc></w:tr></w:tbl>';
  }

  // ── Montar fichas ─────────────────────────────────────────────────────────
  const allImages=[]; var pid=1; const bodyParts=[];

  for (var fi=0;fi<dados.length;fi++) {
    const {m,res}=dados[fi];
    res.sort((a,b)=>a.ano-b.ano||ORDEM.indexOf(a.tipo)-ORDEM.indexOf(b.tipo));
    const labels=res.map(r=>{
      const s=r.tipo==='diagnostica'?('AD '+r.ano):(r.descricao||r.tipo.replace('taf','o TAF')+' '+r.ano);
      return s.replace(/º/g,'o').replace(/ã/g,'a').replace(/ç/g,'c').replace(/õ/g,'o');
    });
    const armaObj=getArma(m.arma);
    const padrao=calcularPadrao(m.arma,m.situacao_funcional);

    if(fi>0) bodyParts.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');

    bodyParts.push(wP((m.posto_graduacao||'')+' '+(armaObj&&armaObj.sigla||'')+' '+(m.nome_guerra||m.nome)+' - FICHA INDIVIDUAL','left','2D4A1E',26,true,0,10));
    bodyParts.push(wP((armaObj&&armaObj.label||'')+' | Padrao '+padrao+' | '+TAF.calcularIdade(m.data_nascimento)+' anos | '+aval.descricao,'left','888888',14,false,0,40));

    // Gráfico global
    bodyParts.push(wP('Evolucao - Mencao Global','left','2D4A1E',18,true,0,16));
    const pngG = gerarPNG(labels, res.map(r=>r.conceito_global), '', 500, 230);
    const ridG = 'rImg'+(allImages.length+1);
    allImages.push({rid:ridG, data:pngG});
    bodyParts.push('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'+wImg(ridG,4800000,2200000,pid++)+'</w:p>');

    // OIIs 2x2
    bodyParts.push(wP('Evolucao por OII','left','2D4A1E',18,true,0,16));
    const oiiCells=[];
    OIIs.forEach(function(o){
      const pngO=gerarPNG(labels,res.map(r=>r[o.campo]),o.label,230,210);
      const ridO='rImg'+(allImages.length+1);
      allImages.push({rid:ridO,data:pngO});
      oiiCells.push('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'+wImg(ridO,2200000,2100000,pid++)+'</w:p>');
    });
    bodyParts.push(wTbl2(oiiCells[0],oiiCells[1]));
    bodyParts.push(wTbl2(oiiCells[2],oiiCells[3]));

    // Tabela histórico
    bodyParts.push(wP('Historico de Avaliacoes','left','2D4A1E',18,true,0,16));
    const cab='<w:tr>'+wTh('Avaliacao',2400)+wTh('Tipo',1000)+wTh('Ano',500)+wTh('Ch',400)+
              wTh('Corrida',700)+wTh('M',500)+wTh('Flexao',700)+wTh('M',500)+
              wTh('Abdom',700)+wTh('M',500)+wTh('Barra',700)+wTh('M',500)+
              wTh('Global',500)+wTh('Suf.',500)+'</w:tr>';
    const linhas=res.map(function(r){
      return '<w:tr>'+
        wTc(r.descricao,2400,'left','FFFFFF','000000',false,14)+
        wTc(labelTipo(r.tipo),1000,'left','FFFFFF','000000',false,14)+
        wTc(r.ano,500,'center','FFFFFF','000000',false,14)+
        wTc((r.chamada||1)+'a',400,'center','FFFFFF','000000',false,14)+
        wTc(r.corrida_12min||'',700,'center','FFFFFF','000000',false,14)+wTcM(r.conceito_corrida)+
        wTc(r.flexao_bracos||'',700,'center','FFFFFF','000000',false,14)+wTcM(r.conceito_flexao_bracos)+
        wTc(r.abdominal_supra||'',700,'center','FFFFFF','000000',false,14)+wTcM(r.conceito_abdominal)+
        wTc(r.flexao_barra||'',700,'center','FFFFFF','000000',false,14)+wTcM(r.conceito_barra)+
        wTcM(r.conceito_global)+wTcM(r.suficiencia)+'</w:tr>';
    }).join('');
    bodyParts.push('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'+
      '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="CCCCCC"/>'+
      '<w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/>'+
      '<w:left w:val="single" w:sz="4" w:color="CCCCCC"/>'+
      '<w:right w:val="single" w:sz="4" w:color="CCCCCC"/>'+
      '<w:insideH w:val="single" w:sz="4" w:color="EDE7D8"/>'+
      '<w:insideV w:val="single" w:sz="4" w:color="EDE7D8"/>'+
      '</w:tblBorders></w:tblPr>'+cab+linhas+'</w:tbl>');
  }

  // ── Montar DOCX ──────────────────────────────────────────────────────────
  const body=bodyParts.join('');
  const imgRels=allImages.map(img=>'<Relationship Id="'+img.rid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/'+img.rid+'.png"/>').join('');

  const docXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'+
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'+
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'+
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'+
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'+
    '<w:body>'+body+
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'+
    '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/>'+
    '</w:sectPr></w:body></w:document>';

  const docRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'+
    imgRels+'</Relationships>';

  const ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
    '<Default Extension="xml" ContentType="application/xml"/>'+
    '<Default Extension="png" ContentType="image/png"/>'+
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'+
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'+
    '</Types>';

  const rr='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'+
    '</Relationships>';

  const styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'+
    '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/>'+
    '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>';

  const zip=new JSZip();
  zip.file('[Content_Types].xml',ct);
  zip.file('_rels/.rels',rr);
  zip.file('word/document.xml',docXml);
  zip.file('word/_rels/document.xml.rels',docRels);
  zip.file('word/styles.xml',styles);
  // Imagens PNG em base64
  allImages.forEach(function(img){
    zip.file('word/media/'+img.rid+'.png', img.data, {base64:true});
  });

  const blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;
  a.download='fichas-individuais-'+aval.descricao.replace(/[^a-zA-Z0-9]/g,'-')+'.docx';
  a.click();
  URL.revokeObjectURL(url);
  toast('Fichas Word geradas para '+dados.length+' militares!');
}


// ═════════════════════════════════════════════════════════════════════════
// MILITARES
// ═════════════════════════════════════════════════════════════════════════
function paginaMilitares() {
  return `
  <div class="page-header">
    <div>
      <div class="page-title">MILITARES CADASTRADOS</div>
      <div class="page-subtitle">Gerencie o cadastro dos militares avaliados</div>
    </div>
    <div class="page-actions">
      <button class="btn btn-ouro" onclick="abrirModalMilitar()">+ Novo Militar</button>
    </div>
  </div>

  <div class="filtro-bar">
    <input id="filtroMil" placeholder="🔍 Buscar por nome..." oninput="filtrarMilitares()">
    <select id="filtroSexo" onchange="filtrarMilitares()">
      <option value="">Todos os sexos</option>
      <option value="M">♂ Masculino</option>
      <option value="F">♀ Feminino</option>
    </select>
    <select id="filtroArma" onchange="filtrarMilitares()">
      <option value="">Todas as armas</option>
      ${ARMAS.map(a => `<option value="${a.value}">${a.sigla} — ${a.label}</option>`).join('')}
    </select>
  </div>

  <div id="tabelaMilitares">
    ${renderTabelaMilitares(state.militares)}
  </div>`;
}

function renderTabelaMilitares(lista) {
  if (!lista.length) return '<div class="empty-state"><div class="empty-icon">👤</div><p>Nenhum militar encontrado.</p></div>';
  return `
  <div class="table-wrap"><table>
    <thead><tr>
      <th>Nome Completo</th><th>Nome de Guerra</th><th>Posto/Grad</th>
      <th>Arma</th><th>LEM</th>
      <th class="center">Sexo</th><th class="center">Idade</th>
      <th class="center">Ações</th>
    </tr></thead>
    <tbody>
      ${lista.map(m => {
        const ao = getArma(m.arma);
        return `<tr>
          <td><strong>${m.nome}</strong></td>
          <td>${m.nome_guerra || '—'}</td>
          <td>${m.posto_graduacao}</td>
          <td><span class="badge badge-B" style="font-size:10px">${ao?.sigla || m.arma || '—'}</span> <span style="font-size:12px">${ao?.label || ''}</span></td>
          <td><span style="font-size:11px;color:var(--texto-leve)">${m.lem || '—'}</span></td>
          <td class="center">${m.sexo === 'M' ? '♂' : '♀'}</td>
          <td class="center">${TAF.calcularIdade(m.data_nascimento)}</td>
          <td class="center" style="white-space:nowrap">
            <button class="btn btn-ghost btn-sm" onclick="abrirModalMilitar(${m.id})">✎</button>
            <button class="btn btn-ghost btn-sm" onclick="navegarComMilitar('ficha',${m.id})" title="Ficha">📋</button>
            <button class="btn btn-danger btn-sm" onclick="excluirMilitar(${m.id})">✕</button>
          </td>
        </tr>`;
      }).join('')}
    </tbody>
  </table></div>`;
}

function filtrarMilitares() {
  const q    = document.getElementById('filtroMil').value.toLowerCase();
  const sex  = document.getElementById('filtroSexo').value;
  const arma = document.getElementById('filtroArma').value;
  const lista = state.militares.filter(m =>
    (!q    || m.nome.toLowerCase().includes(q) || (m.nome_guerra||'').toLowerCase().includes(q)) &&
    (!sex  || m.sexo === sex) &&
    (!arma || m.arma === arma)
  );
  document.getElementById('tabelaMilitares').innerHTML = renderTabelaMilitares(lista);
}

function abrirModalMilitar(id) {
  const m = id ? state.militares.find(x => x.id === id) : null;
  abrirModal(m ? 'Editar Militar' : 'Novo Militar', `
    <div class="form-grid">
      <div class="form-group form-full">
        <label>Nome Completo *</label>
        <input id="mm_nome" placeholder="Nome completo" value="${m?.nome || ''}">
      </div>
      <div class="form-group">
        <label>Nome de Guerra</label>
        <input id="mm_guerra" placeholder="Nome de guerra" value="${m?.nome_guerra || ''}">
      </div>
      <div class="form-group">
        <label>Posto / Graduação *</label>
        <select id="mm_posto">
          ${POSTOS_GRADUACOES.map(p => `<option ${m?.posto_graduacao === p ? 'selected' : ''}>${p}</option>`).join('')}
        </select>
      </div>
    </div>

    <div class="secao-titulo">ARMA / QUADRO / SERVIÇO E SITUAÇÃO FUNCIONAL</div>
    <div class="form-grid">
      <div class="form-group">
        <label>Arma / Quadro / Serviço *</label>
        <select id="mm_arma" onchange="onChangeArma()">
          <option value="">— Selecione —</option>
          ${selectArmas(m?.arma || '')}
        </select>
      </div>
      <div class="form-group">
        <label>Tipo de OM *</label>
        <select id="mm_sitfunc" onchange="onChangeArma()">
          ${SITUACOES_FUNCIONAIS.map(s => `<option value="${s.value}" ${m?.situacao_funcional === s.value ? 'selected' : ''}>${s.label}</option>`).join('')}
        </select>
      </div>
      <!-- LEM e padrão calculados automaticamente -->
      <div class="form-group">
        <label>LEM (automático)</label>
        <div id="mm_lem_display" class="campo-auto">
          ${m?.lem ? `<span class="badge badge-B">${m.lem}</span> — ${lemLabel(m.lem)}` : '<span class="text-muted">Selecione a arma</span>'}
        </div>
        <input type="hidden" id="mm_lem" value="${m?.lem || ''}">
      </div>
      <div class="form-group">
        <label>Padrão exigido (automático)</label>
        <div id="mm_padrao_display" class="campo-auto">
          <span class="text-muted">Selecione a arma e tipo de OM</span>
        </div>
      </div>
    </div>

    <div class="secao-titulo">DADOS PESSOAIS E LOTAÇÃO</div>
    <div class="form-grid">
      <div class="form-group">
        <label>Sexo *</label>
        <select id="mm_sexo">
          <option value="M" ${m?.sexo === 'M' ? 'selected' : ''}>Masculino</option>
          <option value="F" ${m?.sexo === 'F' ? 'selected' : ''}>Feminino</option>
        </select>
      </div>
      <div class="form-group">
        <label>Data de Nascimento *</label>
        <input type="date" id="mm_nasc" value="${m?.data_nascimento || ''}">
      </div>
      <div class="form-group">
        <label>OM</label>
        <input id="mm_om" placeholder="Organização Militar" value="${m?.om || ''}">
      </div>
      <div class="form-group">
        <label>Subunidade</label>
        <input id="mm_su" placeholder="Subunidade" value="${m?.subunidade || ''}">
      </div>
    </div>`,
    `<button class="btn btn-ghost" onclick="fecharModalForce()">Cancelar</button>
     <button class="btn btn-ouro" onclick="salvarMilitar(${id || 0})">💾 Salvar</button>`
  );
  // se editando, disparar atualização do display
  if (m?.arma) setTimeout(onChangeArma, 50);
}

/** Atualiza display de LEM e padrão quando muda arma ou tipo de OM */
function onChangeArma() {
  const armaVal  = document.getElementById('mm_arma')?.value;
  const sitfunc  = document.getElementById('mm_sitfunc')?.value;
  const armaObj  = getArma(armaVal);

  // LEM
  const lemDisplay = document.getElementById('mm_lem_display');
  const lemInput   = document.getElementById('mm_lem');
  if (armaObj && lemDisplay && lemInput) {
    lemInput.value = armaObj.lem;
    lemDisplay.innerHTML = `<span class="badge badge-B">${armaObj.lem}</span> <span style="color:var(--texto-leve);font-size:12px">— ${lemLabel(armaObj.lem)}</span>`;
    if (armaObj.obs) {
      lemDisplay.innerHTML += `<br><span class="text-muted" style="font-size:11px">⚠ ${armaObj.obs}</span>`;
    }
  }

  // Padrão
  const padDisplay = document.getElementById('mm_padrao_display');
  if (armaObj && sitfunc && padDisplay) {
    const padrao = calcularPadrao(armaVal, sitfunc);
    const cores  = { PBD: 'badge-B', PAD: 'badge-MB', PED: 'badge-E' };
    const desc   = { PBD: 'Padrão Básico — mín. Regular', PAD: 'Padrão Avançado — mín. Bom', PED: 'Padrão Especial — mín. Muito Bom' };
    padDisplay.innerHTML = `<span class="badge ${cores[padrao]}">${padrao}</span> <span style="color:var(--texto-leve);font-size:12px">— ${desc[padrao]}</span>`;
  }
}

function lemLabel(lem) {
  return {
    LEMB:  'Linha de Ensino Militar Bélico',
    LEMS:  'Linha de Ensino Militar de Saúde',
    LEMC:  'Linha de Ensino Militar Complementar',
    LEMCT: 'Linha de Ensino Militar Científico-Tecnológico',
  }[lem] || lem;
}

async function salvarMilitar(id) {
  const dados = {
    id: id || null,
    nome:               document.getElementById('mm_nome').value.trim(),
    nome_guerra:        document.getElementById('mm_guerra').value.trim(),
    posto_graduacao:    document.getElementById('mm_posto').value,
    arma:               document.getElementById('mm_arma').value,
    lem:                document.getElementById('mm_lem').value,
    situacao_funcional: document.getElementById('mm_sitfunc').value,
    sexo:               document.getElementById('mm_sexo').value,
    data_nascimento:    document.getElementById('mm_nasc').value,
    om:                 document.getElementById('mm_om').value.trim(),
    subunidade:         document.getElementById('mm_su').value.trim(),
  };
  if (!dados.nome || !dados.data_nascimento) { toast('Nome e data de nascimento são obrigatórios.', true); return; }
  if (!dados.arma) { toast('Selecione a Arma / Quadro / Serviço.', true); return; }
  await window.api.militares.salvar(dados);
  fecharModal();
  toast('Militar salvo com sucesso!');
  await renderizarPagina('militares');
}

async function excluirMilitar(id) {
  if (!confirm('Desativar este militar do sistema?')) return;
  await window.api.militares.excluir(id);
  toast('Militar removido.');
  await renderizarPagina('militares');
}

// ═════════════════════════════════════════════════════════════════════════
// LANÇAR RESULTADOS — tabela inline por TAF
// Fluxo: seleciona o TAF → seleciona a chamada → tabela com todos os
// militares → digita os índices → menção/suficiência aparecem na hora →
// salva tudo de uma vez no final.
// ═════════════════════════════════════════════════════════════════════════

function paginaLancar(avalIdPresel) {
  const avaliacoes = state.avaliacoes;
  const optsAval = avaliacoes.map(a =>
    `<option value="${a.id}" ${a.id == avalIdPresel ? 'selected' : ''}>${a.descricao} (${a.ano})</option>`).join('');

  return `
  <div class="page-header">
    <div>
      <div class="page-title">LANÇAR RESULTADOS</div>
      <div class="page-subtitle">Digite os índices — menção e suficiência calculados automaticamente</div>
    </div>
  </div>

  <div class="card" style="margin-bottom:14px">
    <div class="form-row" style="align-items:flex-end;gap:12px;flex-wrap:wrap">
      <div class="form-group" style="min-width:280px;flex:1">
        <label>Avaliação TAF *</label>
        <select id="taf_sel_aval" onchange="abrirTabelaTAF()">
          <option value="">— Selecione a avaliação —</option>
          ${optsAval}
        </select>
      </div>
      <div class="form-group" style="min-width:160px">
        <label>Chamada *</label>
        <select id="taf_sel_chamada" onchange="abrirTabelaTAF()">
          <option value="1">1ª Chamada</option>
          <option value="2">2ª Chamada</option>
          <option value="3">3ª Chamada</option>
          <option value="4">Chamada Extra</option>
        </select>
      </div>
      <div class="form-group" style="min-width:160px">
        <label>Filtrar por nome</label>
        <input id="taf_filtro" placeholder="🔍 nome..." oninput="filtrarTabelaTAF()">
      </div>
    </div>
  </div>

  <div id="tafTabelaArea">
    <div class="empty-state">
      <div class="empty-icon">📋</div>
      <p>Selecione uma avaliação acima para começar.</p>
    </div>
  </div>`;
}

// Cache das linhas renderizadas (mid → dados)
const _tafLinhas = {};
let _tafAvalAtual = null;
let _tafAvalPendente = null;  // avalId a pré-selecionar ao abrir o lançar       // avaliação selecionada no lançar
let _tafChamadaAtual = 1;       // chamada selecionada
let _tafDataChamadaAtual = null; // data da chamada para cálculo de idade

async function abrirTabelaTAF() {
  const aid     = parseInt(document.getElementById('taf_sel_aval')?.value);
  const chamada = parseInt(document.getElementById('taf_sel_chamada')?.value) || 1;
  if (!aid) return;

  document.getElementById('tafTabelaArea').innerHTML =
    `<div class="card"><p class="text-muted" style="padding:20px">Carregando...</p></div>`;

  // Garantir dados frescos
  await carregarDados();

  if (!state.militares.length) {
    document.getElementById('tafTabelaArea').innerHTML =
      `<div class="empty-state"><div class="empty-icon">👤</div><p>Nenhum militar cadastrado ainda.</p></div>`;
    return;
  }

  const aval = state.avaliacoes.find(a => a.id === aid);
  if (!aval) {
    document.getElementById('tafTabelaArea').innerHTML =
      `<div class="empty-state"><p>Avaliação não encontrada.</p></div>`;
    return;
  }

  const dataChamada = chamada === 1 ? aval.data_1_chamada
                    : chamada === 2 ? aval.data_2_chamada
                    : chamada === 3 ? aval.data_3_chamada
                    : aval.data_chamada_extra;

  _tafAvalAtual        = aval;
  _tafChamadaAtual     = chamada;
  _tafDataChamadaAtual = dataChamada;

  // Buscar resultados desta chamada — sem filtro de ativa para não perder dados
  const todosResultados = await window.api.resultados.listarComMilitares(aid, chamada);

  // Indexar por militar_id
  const mapExist = {};
  todosResultados.forEach(r => {
    if (r.resultado_id != null) mapExist[r.militar_id] = r;
  });

  // Montar _tafLinhas a partir de state.militares
  Object.keys(_tafLinhas).forEach(k => delete _tafLinhas[k]);
  state.militares.forEach(m => {
    const armaObj  = getArma(m.arma);
    const lem      = armaObj?.lem || m.lem || 'LEMB';
    const isLEMB   = lem === 'LEMB';
    const padrao   = calcularPadrao(m.arma, m.situacao_funcional);
    const anoRef   = dataChamada ? dataChamada : (aval.ano ? `${aval.ano}-01-01` : null);
    const idadeRef = TAF.calcularIdade(m.data_nascimento, anoRef);
    const temBarra = isLEMB && idadeRef <= 49;
    const temPPM   = isLEMB && (padrao === 'PAD' || padrao === 'PED') && idadeRef <= 41;
    const ex       = mapExist[m.id] || mapExist[String(m.id)] || {};

    _tafLinhas[m.id] = {
      m, lem, isLEMB, padrao, idadeRef, temBarra, temPPM,
      corrida:  ex.corrida_12min   != null ? String(ex.corrida_12min)   : '',
      flexao:   ex.flexao_bracos   != null ? String(ex.flexao_bracos)   : '',
      abdom:    ex.abdominal_supra != null ? String(ex.abdominal_supra) : '',
      barra:    ex.flexao_barra    != null ? String(ex.flexao_barra)    : '',
      ppm:      ex.ppm_tempo != null ? ex.ppm_tempo : '',
      obs:      ex.observacoes || '',
      tafAlt:   ex.taf_alternativo || 0,
    };
  });

  renderTabelaTAF();
}

function renderTabelaTAF() {
  const filtro = (document.getElementById('taf_filtro')?.value || '').toLowerCase();
  const linhas = Object.values(_tafLinhas).filter(l =>
    !filtro ||
    l.m.nome.toLowerCase().includes(filtro) ||
    (l.m.nome_guerra || '').toLowerCase().includes(filtro)
  );

  if (!linhas.length) {
    document.getElementById('tafTabelaArea').innerHTML =
      `<div class="empty-state"><p>Nenhum militar encontrado.</p></div>`;
    return;
  }

  const coresPad = { PBD: 'badge-B', PAD: 'badge-MB', PED: 'badge-E' };

  const linhasHTML = linhas.map(l => {
    const mid = l.m.id;
    const calcRes = calcularLinhaPreview(l);

    return `
    <tr id="taf-row-${mid}" data-mid="${mid}">
      <td style="white-space:nowrap">
        <strong>${l.m.posto_graduacao}</strong><br>
        <span style="font-size:11px;color:var(--texto-leve)">${l.m.om || ''}</span>
      </td>
      <td>
        <strong>${l.m.nome_guerra || l.m.nome}</strong><br>
        <span style="font-size:11px;color:var(--texto-leve)">${l.m.nome}</span>
      </td>
      <td class="center">${l.m.sexo}</td>
      <td class="center">${l.idadeRef}</td>
      <td class="center"><span class="badge badge-B" style="font-size:10px">${l.lem}</span></td>
      <td class="center"><span class="badge ${coresPad[l.padrao]}" style="font-size:10px">${l.padrao}</span></td>

      <!-- OII: Corrida -->
      <td class="center td-oii">
        <input class="inp-oii" type="number" placeholder="m"
          id="c_${mid}" value="${l.corrida}"
          oninput="onInputOII(${mid})" onkeydown="navOII(event,${mid},'corrida')">
        <div class="oii-badge" id="bc_${mid}">${calcRes.bCorrida}</div>
      </td>

      <!-- OII: Flexão -->
      <td class="center td-oii">
        <input class="inp-oii" type="number" placeholder="rep"
          id="f_${mid}" value="${l.flexao}"
          oninput="onInputOII(${mid})" onkeydown="navOII(event,${mid},'flexao')">
        <div class="oii-badge" id="bf_${mid}">${calcRes.bFlexao}</div>
      </td>

      <!-- OII: Abdominal -->
      <td class="center td-oii">
        <input class="inp-oii" type="number" placeholder="rep"
          id="a_${mid}" value="${l.abdom}"
          oninput="onInputOII(${mid})" onkeydown="navOII(event,${mid},'abdom')">
        <div class="oii-badge" id="ba_${mid}">${calcRes.bAbdom}</div>
      </td>

      <!-- OII: Barra -->
      <td class="center td-oii">
        ${l.temBarra ? `
        <input class="inp-oii" type="number" placeholder="rep"
          id="b_${mid}" value="${l.barra}"
          oninput="onInputOII(${mid})" onkeydown="navOII(event,${mid},'barra')">
        <div class="oii-badge" id="bb_${mid}">${calcRes.bBarra}</div>` : '<span style="color:#ccc">—</span>'}
      </td>

      <!-- OII: PPM -->
      <td class="center td-oii">
        ${l.temPPM ? `
        <input class="inp-oii inp-ppm" type="text" placeholder="MM:SS"
          id="p_${mid}" value="${l.ppm}"
          oninput="onInputOII(${mid})" onkeydown="navOII(event,${mid},'ppm')">
        <div class="oii-badge" id="bp_${mid}">${calcRes.bPPM}</div>` : '<span style="color:#ccc">—</span>'}
      </td>

      <!-- Conceito Global -->
      <td class="center" id="tg_${mid}">${calcRes.global}</td>

      <!-- Suficiência -->
      <td class="center" id="ts_${mid}">${calcRes.sufic}</td>

      <!-- Status salvo -->
      <td class="center" id="tstatus_${mid}">
        ${l.corrida || l.flexao || l.abdom ? '<span class="badge-salvo">✔</span>' : '<span class="badge-pendente">—</span>'}
      </td>
    </tr>`;
  }).join('');

  document.getElementById('tafTabelaArea').innerHTML = `
  <div class="card" style="padding:0;overflow:hidden">
    <div style="display:flex;align-items:center;justify-content:space-between;padding:12px 16px;background:var(--verde-eb)">
      <span style="color:var(--ouro-claro);font-family:'Barlow Condensed',sans-serif;font-size:16px;letter-spacing:1px">
        ${linhas.length} militar(es) — preencha os índices e clique em Salvar Tudo
      </span>
      <div style="display:flex;gap:8px">
        <button class="btn btn-ghost btn-sm" onclick="marcarNR()" style="color:rgba(255,255,255,.7);border-color:rgba(255,255,255,.3)">NR em branco</button>
        <button class="btn btn-ouro" onclick="salvarTodosTAF()">💾 Salvar Tudo</button>
      </div>
    </div>
    <div class="table-wrap">
      <table class="taf-tabela">
        <thead><tr>
          <th>Posto / OM</th>
          <th>Nome</th>
          <th class="center">Sx</th>
          <th class="center">Id</th>
          <th class="center">LEM</th>
          <th class="center">Pad</th>
          <th class="center">Corrida<br><small>metros</small></th>
          <th class="center">Flexão<br><small>reps</small></th>
          <th class="center">Abdom<br><small>reps</small></th>
          <th class="center">Barra<br><small>reps</small></th>
          <th class="center">PPM<br><small>MM:SS</small></th>
          <th class="center">Global</th>
          <th class="center">Sufic</th>
          <th class="center">✔</th>
        </tr></thead>
        <tbody id="tafTbody">${linhasHTML}</tbody>
      </table>
    </div>
    <div style="padding:10px 16px;text-align:right;border-top:1px solid var(--bege-escuro)">
      <button class="btn btn-ouro" onclick="salvarTodosTAF()">💾 Salvar Tudo</button>
    </div>
  </div>`;
}

function filtrarTabelaTAF() {
  renderTabelaTAF();
}

/** Calcula badges de conceito para uma linha */
function calcularLinhaPreview(l) {
  const corrida = parseInt(document.getElementById(`c_${l.m.id}`)?.value) || (l.corrida ? parseInt(l.corrida) : null);
  const flexao  = parseInt(document.getElementById(`f_${l.m.id}`)?.value) || (l.flexao  ? parseInt(l.flexao)  : null);
  const abdom   = parseInt(document.getElementById(`a_${l.m.id}`)?.value) || (l.abdom   ? parseInt(l.abdom)   : null);
  const barra   = parseInt(document.getElementById(`b_${l.m.id}`)?.value) || (l.barra   ? parseInt(l.barra)   : null);
  const ppmStr  = document.getElementById(`p_${l.m.id}`)?.value || l.ppm || null;
  const ppmSeg  = TAF.mmssParaSegundos(ppmStr);

  if (!corrida && !flexao && !abdom) {
    return { bCorrida:'', bFlexao:'', bAbdom:'', bBarra:'', bPPM:'', global:'—', sufic:'—' };
  }

  const res = TAF.calcularTAF({
    dataNascimento:    l.m.data_nascimento,
    dataChamada:       _tafDataChamadaAtual || null,
    anoTAF:            _tafAvalAtual?.ano || null,
    lem:               l.lem,
    sexo:              l.m.sexo,
    situacaoFuncional: l.m.situacao_funcional,
    padrao:            _tafAvalAtual?.padrao || null,
    postoGraduacao:    l.m.posto_graduacao,
    corrida, flexao, abdominal: abdom, barra, ppmTempo: ppmSeg,
    tafAlternativo: false,
  });

  const b = (c) => c ? `<span class="badge badge-${c}" style="font-size:11px">${c}</span>` : '';
  const bs = (c) => c ? `<span class="badge badge-${c}" style="font-size:11px">${c}</span>` : '';

  return {
    bCorrida: b(res.conceitos?.corrida),
    bFlexao:  b(res.conceitos?.flexao),
    bAbdom:   b(res.conceitos?.abdominal),
    bBarra:   b(res.conceitos?.barra),
    bPPM:     bs(res.suficiencias?.ppm),
    global:   b(res.conceitoGlobal) || '—',
    sufic:    bs(res.suficienciaGlobal) || '—',
    res,
  };
}

/** Atualiza badges na linha ao digitar */
function onInputOII(mid) {
  const l = _tafLinhas[mid];
  if (!l) return;
  const calc = calcularLinhaPreview(l);
  const s = (id, v) => { const el = document.getElementById(id); if (el) el.innerHTML = v; };
  s(`bc_${mid}`, calc.bCorrida);
  s(`bf_${mid}`, calc.bFlexao);
  s(`ba_${mid}`, calc.bAbdom);
  s(`bb_${mid}`, calc.bBarra);
  s(`bp_${mid}`, calc.bPPM);
  s(`tg_${mid}`, calc.global || '—');
  s(`ts_${mid}`, calc.sufic  || '—');
  // marca como editado
  const st = document.getElementById(`tstatus_${mid}`);
  if (st) st.innerHTML = '<span class="badge-editado">✎</span>';
}

/** Tab/Enter navega entre campos da linha */
function navOII(event, mid, campo) {
  if (event.key !== 'Tab' && event.key !== 'Enter') return;
  event.preventDefault();
  const l = _tafLinhas[mid];
  if (!l) return;
  const ordemBase = ['corrida', 'flexao', 'abdom'];
  if (l.temBarra) ordemBase.push('barra');
  if (l.temPPM)   ordemBase.push('ppm');
  const idMap = { corrida:`c_${mid}`, flexao:`f_${mid}`, abdom:`a_${mid}`, barra:`b_${mid}`, ppm:`p_${mid}` };
  const idx = ordemBase.indexOf(campo);
  const next = ordemBase[idx + 1];
  if (next) {
    document.getElementById(idMap[next])?.focus();
  } else {
    // vai para o próximo militar
    const mids = Object.keys(_tafLinhas).map(Number);
    const myIdx = mids.indexOf(mid);
    const nextMid = mids[myIdx + 1];
    if (nextMid) document.getElementById(`c_${nextMid}`)?.focus();
  }
}

/** Salva todos os militares visíveis de uma vez */
async function salvarTodosTAF() {
  const aid     = parseInt(document.getElementById('taf_sel_aval')?.value);
  const chamada = _tafChamadaAtual;
  const aval    = _tafAvalAtual;
  const dataChamada = _tafDataChamadaAtual;
  if (!aid) { toast('Selecione a avaliação.', true); return; }

  let salvos = 0, ignorados = 0;

  for (const [midStr, l] of Object.entries(_tafLinhas)) {
    const mid     = parseInt(midStr);
    const corrida = parseInt(document.getElementById(`c_${mid}`)?.value) || null;
    const flexao  = parseInt(document.getElementById(`f_${mid}`)?.value) || null;
    const abdom   = parseInt(document.getElementById(`a_${mid}`)?.value) || null;
    const barra   = parseInt(document.getElementById(`b_${mid}`)?.value) || null;
    const ppmStr  = document.getElementById(`p_${mid}`)?.value || null;

    // Pula linha completamente vazia
    if (!corrida && !flexao && !abdom && !barra && !ppmStr) { ignorados++; continue; }

    const ppmSeg = TAF.mmssParaSegundos(ppmStr);
    const res = TAF.calcularTAF({
      dataNascimento:    l.m.data_nascimento,
      dataChamada:       _tafDataChamadaAtual || null,
      anoTAF:            _tafAvalAtual?.ano || null,
      lem:               l.lem,
      sexo:              l.m.sexo,
      situacaoFuncional: l.m.situacao_funcional,
      padrao:            _tafAvalAtual?.padrao || null,
      postoGraduacao:    l.m.posto_graduacao,
      corrida, flexao, abdominal: abdom, barra, ppmTempo: ppmSeg,
      tafAlternativo: false,
    });

    await window.api.resultados.salvar({
      militar_id: mid, avaliacao_id: aid, chamada,
      corrida_12min: corrida, flexao_bracos: flexao,
      abdominal_supra: abdom, flexao_barra: barra, ppm_tempo: ppmStr,
      conceito_corrida:       res.conceitos?.corrida || null,
      conceito_flexao_bracos: res.conceitos?.flexao  || null,
      conceito_abdominal:     res.conceitos?.abdominal || null,
      conceito_barra:         res.conceitos?.barra   || null,
      conceito_ppm:           res.suficiencias?.ppm  || null,
      conceito_global:        res.conceitoGlobal,
      suficiencia:            res.suficienciaGlobal,
      padrao_verificado:      res.padrao,
      taf_alternativo: 0, observacoes: null,
    });

    // Atualiza ícone de status
    const st = document.getElementById(`tstatus_${mid}`);
    if (st) st.innerHTML = '<span class="badge-salvo">✔</span>';
    salvos++;
  }

  toast(`✔ ${salvos} resultado(s) salvo(s)${ignorados ? ` · ${ignorados} linha(s) em branco ignoradas` : ''}.`);
}

/** Marca todos os em branco como NR no banco */
async function marcarNR() {
  const aid     = parseInt(document.getElementById('taf_sel_aval')?.value);
  const chamada = parseInt(document.getElementById('taf_sel_chamada')?.value) || 1;
  if (!aid) { toast('Selecione a avaliação.', true); return; }
  if (!confirm('Marcar todos os militares SEM resultado como NR (Não Realizado)?')) return;

  let cnt = 0;
  for (const [midStr, l] of Object.entries(_tafLinhas)) {
    const mid = parseInt(midStr);
    const corrida = parseInt(document.getElementById(`c_${mid}`)?.value) || null;
    const flexao  = parseInt(document.getElementById(`f_${mid}`)?.value) || null;
    const abdom   = parseInt(document.getElementById(`a_${mid}`)?.value) || null;
    if (corrida || flexao || abdom) continue; // já tem dado
    await window.api.resultados.salvar({
      militar_id: mid, avaliacao_id: aid, chamada,
      corrida_12min: null, flexao_bracos: null, abdominal_supra: null,
      flexao_barra: null, ppm_tempo: null,
      conceito_corrida: null, conceito_flexao_bracos: null,
      conceito_abdominal: null, conceito_barra: null, conceito_ppm: null,
      conceito_global: 'NR', suficiencia: 'NR', padrao_verificado: l.padrao,
      taf_alternativo: 0, observacoes: 'NR',
    });
    const st = document.getElementById(`tstatus_${mid}`);
    if (st) st.innerHTML = '<span class="badge badge-NR" style="font-size:10px">NR</span>';
    cnt++;
  }
  toast(`${cnt} militar(es) marcados como NR.`);
}


// ═════════════════════════════════════════════════════════════════════════
// RELATÓRIO — substituído por resultados integrados no Gerenciar TAFs
// ═════════════════════════════════════════════════════════════════════════
// (página removida do menu — resultados acessíveis clicando no TAF)

// ═════════════════════════════════════════════════════════════════════════
// FICHA INDIVIDUAL
// ═════════════════════════════════════════════════════════════════════════
// Cache de foto por militar (base64) — persiste na sessão
const _fotosCache = {};

function paginaFicha() {
  return `
  <div class="page-header">
    <div>
      <div class="page-title">FICHA INDIVIDUAL DE DESEMPENHO</div>
      <div class="page-subtitle">Selecione um ou mais militares · Histórico TAF · Gráficos de evolução AD → TAFs</div>
    </div>
    <div class="page-actions" style="display:flex;gap:6px;flex-wrap:wrap">
      <button class="btn btn-ghost" onclick="selecionarTodosMilitares()">☑ Selecionar todos</button>
      <button class="btn btn-ghost" onclick="limparSelecaoMilitares()">☐ Limpar</button>
      <button class="btn btn-ouro" onclick="exportarFichasSelecionadasPDF()">📄 PDF Fichas</button>
      <button class="btn btn-ghost" onclick="exportarFichasSelecionadasWord()">📝 Word Fichas</button>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:280px 1fr;gap:16px;height:calc(100vh - 140px)">

    <!-- Lista de militares -->
    <div class="card" style="margin-bottom:0;overflow:hidden;display:flex;flex-direction:column">
      <div class="card-title" style="flex-shrink:0">MILITARES (${state.militares.length})</div>
      <input id="fichaFiltroNome" placeholder="🔍 Filtrar..." oninput="filtrarListaFicha()"
        style="margin:0 0 8px;height:28px;font-size:12px">
      <div id="listaMilitaresFicha" style="overflow-y:auto;flex:1">
        ${renderListaMilitaresFicha(state.militares)}
      </div>
    </div>

    <!-- Área de conteúdo da ficha -->
    <div style="overflow-y:auto">
      <div id="fichaArea">
        <div class="empty-state"><div class="empty-icon">👤</div>
          <p>Selecione um militar ao lado para ver sua ficha.<br>
          Selecione múltiplos para baixar em lote.</p>
        </div>
      </div>
    </div>
  </div>`;
}

function renderListaMilitaresFicha(lista) {
  return lista.map(m => `
    <div class="ficha-mil-item" id="fichaItem_${m.id}"
         onclick="selecionarMilitarFicha(${m.id})"
         style="display:flex;align-items:center;gap:8px;padding:7px 10px;cursor:pointer;border-radius:4px;margin-bottom:2px;border:1px solid transparent">
      <input type="checkbox" id="fichaChk_${m.id}" onclick="event.stopPropagation()"
             onchange="onCheckMilitarFicha(${m.id},this.checked)"
             style="flex-shrink:0;cursor:pointer">
      <div style="flex:1;min-width:0">
        <div style="font-size:12px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
          ${m.nome_guerra || m.nome}
        </div>
        <div style="font-size:10px;color:var(--texto-leve)">${m.posto_graduacao} · ${m.arma||'—'}</div>
      </div>
    </div>`).join('');
}

let _fichasMilSelecionados = new Set();

function filtrarListaFicha() {
  const q = document.getElementById('fichaFiltroNome').value.toLowerCase();
  const lista = q ? state.militares.filter(m =>
    (m.nome_guerra||m.nome||'').toLowerCase().includes(q) ||
    (m.posto_graduacao||'').toLowerCase().includes(q)
  ) : state.militares;
  document.getElementById('listaMilitaresFicha').innerHTML = renderListaMilitaresFicha(lista);
  // Restaurar checkboxes
  _fichasMilSelecionados.forEach(id => {
    const chk = document.getElementById(`fichaChk_${id}`);
    if (chk) chk.checked = true;
  });
}

function selecionarTodosMilitares() {
  state.militares.forEach(m => {
    _fichasMilSelecionados.add(m.id);
    const chk = document.getElementById(`fichaChk_${m.id}`);
    if (chk) chk.checked = true;
  });
}

function limparSelecaoMilitares() {
  _fichasMilSelecionados.clear();
  document.querySelectorAll('[id^="fichaChk_"]').forEach(c => c.checked = false);
  document.querySelectorAll('.ficha-mil-item').forEach(el => {
    el.style.background = ''; el.style.borderColor = 'transparent';
  });
  document.getElementById('fichaArea').innerHTML =
    '<div class="empty-state"><div class="empty-icon">👤</div><p>Selecione um militar ao lado.</p></div>';
}

function onCheckMilitarFicha(mid, checked) {
  if (checked) _fichasMilSelecionados.add(mid);
  else _fichasMilSelecionados.delete(mid);
  const item = document.getElementById(`fichaItem_${mid}`);
  if (item) {
    item.style.background = checked ? 'rgba(45,74,30,.08)' : '';
    item.style.borderColor = checked ? 'var(--verde-claro)' : 'transparent';
  }
}

async function selecionarMilitarFicha(mid) {
  // Clique na linha → seleciona e carrega ficha
  const chk = document.getElementById(`fichaChk_${mid}`);
  const jaSelecionado = _fichasMilSelecionados.has(mid);
  // Se já estava selecionado e é o único, deschecar; senão só mudar o foco
  chk.checked = true;
  onCheckMilitarFicha(mid, true);
  await carregarFichaMilitar(mid);
}

async function exportarFichasSelecionadasPDF() {
  if (!_fichasMilSelecionados.size) {
    if (!_fichaMilAtual) { toast('Selecione ao menos um militar.', true); return; }
    _fichasMilSelecionados.add(_fichaMilAtual);
  }
  const ids = [..._fichasMilSelecionados];
  await _gerarFichasPDF(ids);
}

async function exportarFichasSelecionadasWord() {
  if (!_fichasMilSelecionados.size) {
    if (!_fichaMilAtual) { toast('Selecione ao menos um militar.', true); return; }
    _fichasMilSelecionados.add(_fichaMilAtual);
  }
  const ids = [..._fichasMilSelecionados];
  // Montar lista de dados
  const ORDEM_TIPO = ['diagnostica','1taf','2taf','3taf'];
  const dadosMilitares = await Promise.all(ids.map(async mid => {
    const m   = state.militares.find(x => x.id === mid);
    const res = await window.api.resultados.porMilitar(mid);
    return {m, res};
  }));
  dadosMilitares.sort((a,b)=>(a.m.nome_guerra||a.m.nome).localeCompare(b.m.nome_guerra||b.m.nome));
  await _gerarFichasWord(dadosMilitares);
}

async function _gerarFichasPDF(ids) {
  toast(`Gerando ${ids.length} ficha(s) PDF...`);
  const ORDEM_TIPO = ['diagnostica','1taf','2taf','3taf'];
  const dadosMilitares = await Promise.all(ids.map(async mid => {
    const m   = state.militares.find(x => x.id === mid);
    const res = await window.api.resultados.porMilitar(mid);
    return {m, res};
  }));
  dadosMilitares.sort((a,b)=>(a.m.nome_guerra||a.m.nome).localeCompare(b.m.nome_guerra||b.m.nome));

  const fichasHTML = dadosMilitares.map(({m, res}) => {
    res.sort((a,b)=>a.ano-b.ano||ORDEM_TIPO.indexOf(a.tipo)-ORDEM_TIPO.indexOf(b.tipo));
    const labels = res.map(r => r.tipo==='diagnostica'
      ? `AD ${r.ano}` : (r.descricao||r.tipo.replace('taf','ºTAF')+' '+r.ano));
    const armaObj=getArma(m.arma), padrao=calcularPadrao(m.arma,m.situacao_funcional);
    const fotoSrc=_fotosCache[m.id]||null;
    const OIIs=[
      {campo:'conceito_barra',label:'Barra Fixa'},
      {campo:'conceito_flexao_bracos',label:'Flexão de Braço'},
      {campo:'conceito_abdominal',label:'Abdominal'},
      {campo:'conceito_corrida',label:'Corrida 12min'},
    ];
    const svgGlobal=res.length>0
      ?svgLinhaEvolucao(labels,res.map(r=>r.conceito_global),740,140)
      :'<p style="color:#aaa;text-align:center;padding:10px">Sem dados</p>';
    const svgsOII=OIIs.map(o=>svgLinhaEvolucao(labels,res.map(r=>r[o.campo]),230,120));
    const linhas=res.map(r=>`<tr>
      <td>${r.descricao}</td><td>${r.ano}</td><td>${r.chamada}ª</td>
      <td>${r.corrida_12min||'—'}</td><td class="m">${r.conceito_corrida||''}</td>
      <td>${r.flexao_bracos||'—'}</td><td class="m">${r.conceito_flexao_bracos||''}</td>
      <td>${r.abdominal_supra||'—'}</td><td class="m">${r.conceito_abdominal||''}</td>
      <td>${r.flexao_barra||'—'}</td><td class="m">${r.conceito_barra||''}</td>
      <td class="g">${r.conceito_global||'NR'}</td><td>${r.suficiencia||'NR'}</td>
    </tr>`).join('');
    return `<div class="ficha">
      <div class="topo">
        ${fotoSrc?`<img class="foto" src="${fotoSrc}">`:'<div class="foto-ph">👤</div>'}
        <div class="dados">
          <div class="nome">${m.posto_graduacao} ${armaObj?.sigla||''} ${m.nome_guerra||m.nome}</div>
          <div class="sub">${m.nome}</div>
          <div class="dgrid">
            <span><b>Nasc.:</b> ${formatarData(m.data_nascimento)}</span>
            <span><b>Sexo:</b> ${m.sexo==='M'?'Masculino':'Feminino'}</span>
            <span><b>Arma:</b> ${armaObj?.label||m.arma||'—'}</span>
            <span><b>LEM:</b> ${m.lem||'—'}</span>
            <span><b>Padrão:</b> ${padrao}</span>
            <span><b>OM:</b> ${m.om||'—'}</span>
          </div>
        </div>
      </div>
      <div class="sec">Evolução — Menção Global</div>
      <div class="graf-global">${svgGlobal}</div>
      <div class="sec">Evolução por OII</div>
      <div class="oii-grid">
        ${OIIs.map((o,i)=>`<div class="oii-item"><div class="oii-lbl">${o.label}</div>${svgsOII[i]}</div>`).join('')}
      </div>
      <div class="sec">Histórico</div>
      <table>
        <thead><tr><th>Avaliação</th><th>Ano</th><th>Ch</th>
          <th>Corrida</th><th>M</th><th>Flexão</th><th>M</th>
          <th>Abdom</th><th>M</th><th>Barra</th><th>M</th>
          <th>Global</th><th>Suf</th></tr></thead>
        <tbody>${linhas}</tbody>
      </table>
    </div>`;
  }).join('');

  const win=window.open('','_blank');
  if(!win){toast('Permita popups para PDF.',true);return;}
  win.document.write(`<!DOCTYPE html><html lang="pt-BR"><head>
  <meta charset="UTF-8"><title>Fichas</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@400;500&display=swap');
    @page{size:A4 landscape;margin:7mm 9mm;}
    *{margin:0;padding:0;box-sizing:border-box;}
    body{font-family:'Barlow',sans-serif;font-size:9px;color:#1a2010;}
    .ficha{page-break-after:always;page-break-inside:avoid;}.ficha:last-child{page-break-after:auto;}
    .topo{display:flex;align-items:flex-start;gap:10px;border-bottom:2px solid #c9a227;padding-bottom:7px;margin-bottom:7px;}
    .foto{width:60px;height:76px;object-fit:cover;border:1px solid #bbb;border-radius:2px;flex-shrink:0;}
    .foto-ph{width:60px;height:76px;background:#f0ede6;border:1px dashed #ccc;border-radius:2px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:22px;color:#ccc;}
    .dados{flex:1;}.nome{font-family:'Barlow Condensed',sans-serif;font-size:15px;color:#2d4a1e;font-weight:700;}
    .sub{font-size:9px;color:#888;margin-bottom:4px;}.dgrid{display:flex;flex-wrap:wrap;gap:2px 14px;}
    .sec{font-family:'Barlow Condensed',sans-serif;font-size:9px;color:#2d4a1e;border-bottom:1px solid #e0d8c8;padding-bottom:2px;margin:5px 0 3px;text-transform:uppercase;}
    .graf-global svg,.oii-item svg{width:100%;height:auto;display:block;}
    .oii-grid{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:5px;margin-bottom:5px;}
    .oii-lbl{font-size:8px;color:#888;margin-bottom:1px;text-align:center;}
    table{width:100%;border-collapse:collapse;font-size:8px;}
    th{background:#2d4a1e;color:#c9a227;padding:2px 4px;font-size:8px;}
    td{padding:2px 4px;border-bottom:1px solid #ede7d8;}tr:nth-child(even) td{background:#f9f6f0;}
    td.m{font-weight:700;text-align:center;}td.g{font-weight:700;text-align:center;}
  </style></head><body>
  ${fichasHTML}
  <script>window.onload=function(){setTimeout(function(){window.print();},500);};</script>
  </body></html>`);
  win.document.close();
}

async function _gerarFichasWord(dadosMilitares) {
  if (typeof JSZip==='undefined'){toast('JSZip nao carregado.',true);return;}
  toast('Gerando Word para '+dadosMilitares.length+' ficha(s)...');

  const ORDEM=['diagnostica','1taf','2taf','3taf'];
  const MVAL={E:5,MB:4,B:3,R:2,I:1,NR:0};
  const MCOR={E:'#1565C0',MB:'#2E7D32',B:'#558B2F',R:'#F9A825',I:'#C62828',NR:'#9E9E9E'};
  const YLABELS=['NR','I','R','B','MB','E'];
  const OBG={E:'D4EDDA',MB:'CCE5FF',B:'D4EDDA',R:'FFF3CD',I:'F8D7DA',S:'D4EDDA',NS:'F8D7DA',NR:'E2E3E5','':'FFFFFF'};
  const OFG={E:'155724',MB:'1A4A6E',B:'2D4A1E',R:'856404',I:'721C24',S:'155724',NS:'721C24',NR:'383D41','':'000000'};
  const OIIs=[
    {campo:'conceito_barra',        label:'Barra Fixa'},
    {campo:'conceito_flexao_bracos',label:'Flexao de Braco'},
    {campo:'conceito_abdominal',    label:'Abdominal'},
    {campo:'conceito_corrida',      label:'Corrida 12min'},
  ];

  function gerarPNG(labels,mencoes,titulo,W,H){
    const canvas=document.createElement('canvas');
    const DPR=2; canvas.width=W*DPR; canvas.height=H*DPR;
    const ctx=canvas.getContext('2d'); ctx.scale(DPR,DPR);
    const PAD={top:28,right:16,bottom:36,left:40};
    const pw=W-PAD.left-PAD.right, ph=H-PAD.top-PAD.bottom;
    ctx.fillStyle='#ffffff'; ctx.fillRect(0,0,W,H);
    if(titulo){ctx.fillStyle='#888888';ctx.font='bold 10px Calibri,Arial';ctx.textAlign='center';ctx.fillText(titulo,W/2,14);}
    ctx.lineWidth=0.8;
    for(let i=0;i<=5;i++){
      const y=PAD.top+ph-(i/5)*ph;
      ctx.strokeStyle='#CCCCCC'; ctx.beginPath(); ctx.moveTo(PAD.left,y); ctx.lineTo(PAD.left+pw,y); ctx.stroke();
      ctx.fillStyle='#333333'; ctx.font='bold 10px Calibri,Arial'; ctx.textAlign='right';
      ctx.fillText(YLABELS[i],PAD.left-4,y+4);
    }
    ctx.strokeStyle='#CCCCCC'; ctx.lineWidth=1;
    ctx.beginPath(); ctx.moveTo(PAD.left,PAD.top+ph); ctx.lineTo(PAD.left+pw,PAD.top+ph); ctx.stroke();
    const n=labels.length;
    const xPos=i=>PAD.left+(n===1?pw/2:(i/(n-1))*pw);
    const yPos=v=>PAD.top+ph-(v/5)*ph;
    const vals=mencoes.map(m=>MVAL[m]||0);
    if(n>1){
      ctx.strokeStyle='#2D4A1E'; ctx.lineWidth=2.5;
      ctx.beginPath(); ctx.moveTo(xPos(0),yPos(vals[0]));
      for(let i=1;i<n;i++) ctx.lineTo(xPos(i),yPos(vals[i])); ctx.stroke();
    }
    mencoes.forEach((m,i)=>{
      const cor=MCOR[m]||'#607D8B', x=xPos(i), y=yPos(vals[i]);
      ctx.beginPath(); ctx.arc(x,y,7,0,Math.PI*2); ctx.fillStyle=cor; ctx.fill();
      ctx.fillStyle=cor; ctx.font='bold 11px Calibri,Arial'; ctx.textAlign='center';
      ctx.fillText(m,x,y-11);
    });
    labels.forEach((lbl,i)=>{
      ctx.fillStyle='#222222'; ctx.font='bold 10px Calibri,Arial'; ctx.textAlign='center';
      ctx.fillText(lbl,xPos(i),PAD.top+ph+16);
    });
    return canvas.toDataURL('image/png').split(',')[1];
  }

  function esc(v){return String(v||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
  function wTc(t,w,jc,bg,fg,b,sz){
    return '<w:tc><w:tcPr><w:tcW w:w="'+w+'" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="'+bg+'"/></w:tcPr>'+
      '<w:p><w:pPr><w:jc w:val="'+jc+'"/><w:spacing w:before="0" w:after="0"/></w:pPr>'+
      '<w:r><w:rPr>'+(b?'<w:b/>':'')+'<w:color w:val="'+fg+'"/><w:sz w:val="'+sz+'"/></w:rPr>'+
      '<w:t xml:space="preserve">'+esc(t)+'</w:t></w:r></w:p></w:tc>';
  }
  function wTh(t,w){return wTc(t,w,'center','2D4A1E','C9A227',true,16);}
  function wTcM(v){return wTc(v||'',500,'center',OBG[v||'']||'FFFFFF',OFG[v||'']||'000000',true,15);}
  function wP(t,jc,color,sz,b,bef,aft){
    return '<w:p><w:pPr><w:jc w:val="'+jc+'"/><w:spacing w:before="'+(bef||0)+'" w:after="'+(aft||60)+'"/></w:pPr>'+
      '<w:r><w:rPr>'+(b?'<w:b/>':'')+'<w:color w:val="'+color+'"/><w:sz w:val="'+sz+'"/></w:rPr>'+
      '<w:t xml:space="preserve">'+esc(t)+'</w:t></w:r></w:p>';
  }
  function wImg(rid,cx,cy,pid){
    return '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'+
      '<wp:inline distT="0" distB="0" distL="0" distR="0">'+
      '<wp:extent cx="'+cx+'" cy="'+cy+'"/>'+
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>'+
      '<wp:docPr id="'+pid+'" name="Img'+pid+'"/>'+
      '<wp:cNvGraphicFramePr/>'+
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'+
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'+
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'+
      '<pic:nvPicPr><pic:cNvPr id="'+pid+'" name="Img'+pid+'"/><pic:cNvPicPr/></pic:nvPicPr>'+
      '<pic:blipFill><a:blip r:embed="'+rid+'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'+
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'+
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm>'+
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'+
      '</pic:pic></a:graphicData></a:graphic>'+
      '</wp:inline></w:drawing></w:r>';
  }
  function wTbl2(a,b){
    return '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'+
      '<w:tblBorders><w:top w:val="none"/><w:bottom w:val="none"/>'+
      '<w:left w:val="none"/><w:right w:val="none"/>'+
      '<w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>'+
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr>'+a+'</w:tc>'+
      '<w:tc><w:tcPr><w:tcW w:w="50" w:type="pct"/></w:tcPr>'+b+'</w:tc></w:tr></w:tbl>';
  }

  const allImages=[]; var pid=1; const bodyParts=[];

  for(var fi=0;fi<dadosMilitares.length;fi++){
    const {m,res}=dadosMilitares[fi];
    res.sort((a,b)=>a.ano-b.ano||ORDEM.indexOf(a.tipo)-ORDEM.indexOf(b.tipo));
    const labels=res.map(r=>{
      const s=r.tipo==='diagnostica'?('AD '+r.ano):(r.descricao||r.tipo.replace('taf','o TAF')+' '+r.ano);
      return s.replace(/º/g,'o').replace(/ã/g,'a').replace(/ç/g,'c').replace(/õ/g,'o');
    });
    const armaObj=getArma(m.arma), padrao=calcularPadrao(m.arma,m.situacao_funcional);
    if(fi>0) bodyParts.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
    bodyParts.push(wP((m.posto_graduacao||'')+' '+(armaObj&&armaObj.sigla||'')+' '+(m.nome_guerra||m.nome)+' - FICHA INDIVIDUAL','left','2D4A1E',26,true,0,10));
    bodyParts.push(wP((armaObj&&armaObj.label||'')+' | Padrao '+padrao+' | '+TAF.calcularIdade(m.data_nascimento)+' anos','left','888888',14,false,0,40));
    bodyParts.push(wP('Evolucao - Mencao Global','left','2D4A1E',18,true,0,16));
    const pngG=gerarPNG(labels,res.map(r=>r.conceito_global),'',500,230);
    const ridG='rImg'+(allImages.length+1); allImages.push({rid:ridG,data:pngG});
    bodyParts.push('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'+wImg(ridG,4800000,2200000,pid++)+'</w:p>');
    bodyParts.push(wP('Evolucao por OII','left','2D4A1E',18,true,0,16));
    const oiiCells=[];
    OIIs.forEach(function(o){
      const pngO=gerarPNG(labels,res.map(r=>r[o.campo]),o.label,230,210);
      const ridO='rImg'+(allImages.length+1); allImages.push({rid:ridO,data:pngO});
      oiiCells.push('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'+wImg(ridO,2200000,2100000,pid++)+'</w:p>');
    });
    bodyParts.push(wTbl2(oiiCells[0],oiiCells[1]));
    bodyParts.push(wTbl2(oiiCells[2],oiiCells[3]));
    bodyParts.push(wP('Historico de Avaliacoes','left','2D4A1E',18,true,0,16));
    const cab='<w:tr>'+wTh('Avaliacao',2400)+wTh('Tipo',1000)+wTh('Ano',500)+wTh('Ch',400)+
      wTh('Corrida',700)+wTh('M',500)+wTh('Flexao',700)+wTh('M',500)+
      wTh('Abdom',700)+wTh('M',500)+wTh('Barra',700)+wTh('M',500)+
      wTh('Global',500)+wTh('Suf.',500)+'</w:tr>';
    const linhas=res.map(function(r){
      return '<w:tr>'+wTc(r.descricao,2400,'left','FFFFFF','000000',false,14)+
        wTc(labelTipo(r.tipo),1000,'left','FFFFFF','000000',false,14)+
        wTc(r.ano,500,'center','FFFFFF','000000',false,14)+
        wTc((r.chamada||1)+'a',400,'center','FFFFFF','000000',false,14)+
        wTc(r.corrida_12min||'',700,'center','FFFFFF','000000',false,14)+wTcM(r.conceito_corrida)+
        wTc(r.flexao_bracos||'',700,'center','FFFFFF','000000',false,14)+wTcM(r.conceito_flexao_bracos)+
        wTc(r.abdominal_supra||'',700,'center','FFFFFF','000000',false,14)+wTcM(r.conceito_abdominal)+
        wTc(r.flexao_barra||'',700,'center','FFFFFF','000000',false,14)+wTcM(r.conceito_barra)+
        wTcM(r.conceito_global)+wTcM(r.suficiencia)+'</w:tr>';
    }).join('');
    bodyParts.push('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'+
      '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="CCCCCC"/>'+
      '<w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/>'+
      '<w:left w:val="single" w:sz="4" w:color="CCCCCC"/>'+
      '<w:right w:val="single" w:sz="4" w:color="CCCCCC"/>'+
      '<w:insideH w:val="single" w:sz="4" w:color="EDE7D8"/>'+
      '<w:insideV w:val="single" w:sz="4" w:color="EDE7D8"/>'+
      '</w:tblBorders></w:tblPr>'+cab+linhas+'</w:tbl>');
  }

  const body=bodyParts.join('');
  const imgRels=allImages.map(img=>'<Relationship Id="'+img.rid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/'+img.rid+'.png"/>').join('');
  const docXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'+
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'+
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'+
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'+
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'+
    '<w:body>'+body+
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'+
    '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/>'+
    '</w:sectPr></w:body></w:document>';
  const docRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'+
    imgRels+'</Relationships>';
  const ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
    '<Default Extension="xml" ContentType="application/xml"/>'+
    '<Default Extension="png" ContentType="image/png"/>'+
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'+
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'+
    '</Types>';
  const rr='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'+
    '</Relationships>';
  const styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>'+
    '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/>'+
    '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>';

  const zip=new JSZip();
  zip.file('[Content_Types].xml',ct); zip.file('_rels/.rels',rr);
  zip.file('word/document.xml',docXml); zip.file('word/_rels/document.xml.rels',docRels);
  zip.file('word/styles.xml',styles);
  allImages.forEach(function(img){ zip.file('word/media/'+img.rid+'.png',img.data,{base64:true}); });
  const blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download='fichas-individuais.docx'; a.click();
  URL.revokeObjectURL(url);
  toast('Fichas Word geradas para '+dadosMilitares.length+' militar(es)!');
}

async function carregarFichaMilitar(mid) {
  if (!mid) return;
  _fichaMilAtual = mid;
  const m = state.militares.find(x => x.id === mid);
  const resultados = await window.api.resultados.porMilitar(mid);
  const armaObj = getArma(m.arma);
  const padrao  = calcularPadrao(m.arma, m.situacao_funcional);
  const cPad    = { PBD:'badge-B', PAD:'badge-MB', PED:'badge-E' };

  // Ordenar: AD primeiro, depois 1ºTAF, 2ºTAF etc por ano
  const ORDEM_TIPO = ['diagnostica','1taf','2taf','3taf'];
  resultados.sort((a,b) => a.ano - b.ano || ORDEM_TIPO.indexOf(a.tipo) - ORDEM_TIPO.indexOf(b.tipo));

  // Label do eixo X: usa descrição real da avaliação (ex: "AD 2026", "1º TAF 2026")
  // e dataChamada correta para cálculo de idade
  const pontos = resultados.map(r => {
    // Data da chamada correspondente ao número da chamada
    const dataChamada = r.chamada === 1 ? r.data_1_chamada
                      : r.chamada === 2 ? r.data_2_chamada
                      : r.chamada === 3 ? r.data_3_chamada
                      : r.data_chamada_extra;

    // Label curto para eixo X
    // Label: AD ou descrição real ("1º TAF 2026")
    const labelCurto = r.tipo === 'diagnostica'
      ? `AD ${r.ano}`
      : (r.descricao || (r.tipo.replace('taf','ºTAF') + ' ' + r.ano));

    return {
      label:      labelCurto,
      labelFull:  r.descricao || labelCurto,
      ano:        r.ano,
      chamada:    r.chamada,
      dataChamada,
      corrida:   r.corrida_12min,
      flexao:    r.flexao_bracos,
      abdominal: r.abdominal_supra,
      barra:     r.flexao_barra,
      global:    r.conceito_global,
      suf:       r.suficiencia,
      c_corrida: r.conceito_corrida,
      c_flexao:  r.conceito_flexao_bracos,
      c_abd:     r.conceito_abdominal,
      c_barra:   r.conceito_barra,
      // Idade na data da chamada
      idadeNaChamada: TAF.calcularIdade(m.data_nascimento, dataChamada),
    };
  });

  // Idade HOJE (para exibição na ficha)
  const idadeHoje = TAF.calcularIdade(m.data_nascimento);

  const fotoSrc = _fotosCache[mid] || null;

  document.getElementById('fichaArea').innerHTML = `
  <div id="fichaConteudo">

    <!-- Cabeçalho da ficha -->
    <div class="card ficha-header-card">
      <div class="ficha-header">
        <!-- Foto -->
        <div class="ficha-foto-area">
          <div id="fichaFotoBox" class="ficha-foto-box"
               onclick="document.getElementById('fichaFotoInput').click()"
               title="Clique para adicionar foto ou arraste uma imagem aqui"
               ondragover="event.preventDefault();this.classList.add('foto-drag')"
               ondragleave="this.classList.remove('foto-drag')"
               ondrop="onDropFoto(event,${mid})">
            ${fotoSrc
              ? `<img src="${fotoSrc}" class="ficha-foto-img" id="fichaFotoImg">`
              : `<div class="ficha-foto-placeholder">
                  <div style="font-size:32px;color:#ccc">👤</div>
                  <div style="font-size:10px;color:#aaa;margin-top:4px">Clique ou arraste<br>a foto aqui</div>
                </div>`}
          </div>
          <input type="file" id="fichaFotoInput" accept="image/*" style="display:none"
                 onchange="onSelecionarFoto(event,${mid})">
          ${fotoSrc ? `<button class="btn btn-danger btn-sm" style="margin-top:4px;width:100%" onclick="removerFoto(${mid})">✕ Remover</button>` : ''}
        </div>

        <!-- Dados do militar -->
        <div class="ficha-dados">
          <div class="ficha-nome">${m.posto_graduacao} ${m.nome_guerra || m.nome}</div>
          <div class="ficha-subinfo">${m.nome}</div>
          <div class="ficha-grid">
            <div><span class="ficha-label">Data Nasc.</span><strong>${formatarData(m.data_nascimento)}</strong></div>
            <div><span class="ficha-label">Idade</span><strong>${idadeHoje} anos</strong></div>
            <div><span class="ficha-label">Sexo</span><strong>${m.sexo === 'M' ? 'Masculino' : 'Feminino'}</strong></div>
            <div><span class="ficha-label">Arma</span><strong>${armaObj?.sigla || m.arma || '—'} — ${armaObj?.label || ''}</strong></div>
            <div><span class="ficha-label">LEM</span><strong>${m.lem || '—'}</strong></div>
            <div><span class="ficha-label">Padrão</span><strong><span class="badge ${cPad[padrao]}">${padrao}</span></strong></div>
            <div><span class="ficha-label">OM</span><strong>${m.om || '—'}</strong></div>
            <div><span class="ficha-label">Subunidade</span><strong>${m.subunidade || '—'}</strong></div>
          </div>
        </div>
      </div>
    </div>

    ${resultados.length === 0
      ? '<div class="card empty-state"><p>Nenhum resultado registrado para este militar.</p></div>'
      : `
    <!-- Gráfico de evolução geral (Média/Global) -->
    <div class="card">
      <div class="card-title">📈 EVOLUÇÃO — MENÇÃO GLOBAL</div>
      <canvas id="grafEvolGlobal" style="width:100%;height:200px"></canvas>
    </div>

    <!-- Gráficos por OII — 2x2 -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:16px">
      <div class="card" style="margin-bottom:0">
        <div class="card-title" style="font-size:13px">Barra Fixa</div>
        <canvas id="grafBarra" style="width:100%;height:160px"></canvas>
      </div>
      <div class="card" style="margin-bottom:0">
        <div class="card-title" style="font-size:13px">Flexão de Braço</div>
        <canvas id="grafFlexao" style="width:100%;height:160px"></canvas>
      </div>
      <div class="card" style="margin-bottom:0">
        <div class="card-title" style="font-size:13px">Abdominal</div>
        <canvas id="grafAbdom" style="width:100%;height:160px"></canvas>
      </div>
      <div class="card" style="margin-bottom:0">
        <div class="card-title" style="font-size:13px">Corrida 12 min</div>
        <canvas id="grafCorrida" style="width:100%;height:160px"></canvas>
      </div>
    </div>

    <!-- Tabela histórica -->
    <div class="card">
      <div class="card-title">HISTÓRICO DE AVALIAÇÕES</div>
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th>Avaliação</th><th>Tipo</th><th>Ano</th><th class="center">Cham.</th>
            <th class="center">Corrida</th><th class="center">Flexão</th>
            <th class="center">Abdom</th><th class="center">Barra</th>
            <th class="center">PPM</th><th class="center">Global</th><th class="center">Sufic.</th>
          </tr></thead>
          <tbody>
            ${resultados.map(r => `
            <tr>
              <td>${r.descricao}</td>
              <td>${labelTipo(r.tipo)}</td>
              <td>${r.ano}</td>
              <td class="center">${r.chamada}ª</td>
              <td class="center">${r.corrida_12min||'—'} ${badgeConceito(r.conceito_corrida)}</td>
              <td class="center">${r.flexao_bracos||'—'} ${badgeConceito(r.conceito_flexao_bracos)}</td>
              <td class="center">${r.abdominal_supra||'—'} ${badgeConceito(r.conceito_abdominal)}</td>
              <td class="center">${r.flexao_barra||'—'} ${badgeConceito(r.conceito_barra)}</td>
              <td class="center">${r.ppm_tempo||'—'} ${badgeConceito(r.conceito_ppm)}</td>
              <td class="center">${badgeConceito(r.conceito_global)}</td>
              <td class="center">${badgeSuficiencia(r.suficiencia)}</td>
            </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`}
  </div>`;

  // Renderizar gráficos de evolução — múltiplas tentativas para garantir offsetWidth
  if (resultados.length > 0) {
    const _tentarRenderizar = (tentativa) => {
      const c1 = document.getElementById('grafEvolGlobal');
      if (!c1 || c1.offsetWidth === 0) {
        if (tentativa < 10) setTimeout(() => _tentarRenderizar(tentativa+1), 80);
        return;
      }
      renderGraficosEvolucao(pontos);
    };
    setTimeout(() => _tentarRenderizar(0), 100);
  }
}

// ── Gráficos de evolução (linha) ──────────────────────────────────────────
const MENCAO_VAL = { E:5, MB:4, B:3, R:2, I:1, NR:0 };
const MENCAO_LABEL = ['','I','R','B','MB','E'];
const COR_MENCAO = { E:'#1565C0', MB:'#2e7d32', B:'#f9a825', R:'#e65100', I:'#c62828', NR:'#9e9e9e' };

function renderGraficosEvolucao(pontos) {
  const labels = pontos.map(p => p.label);

  // Gráfico global
  desenharLinhaEvolucao('grafEvolGlobal', labels,
    pontos.map(p => MENCAO_VAL[p.global] || 0),
    pontos.map(p => p.global),
    '#2d4a1e', 'Menção Global'
  );

  // OIIs
  desenharLinhaEvolucao('grafBarra', labels,
    pontos.map(p => MENCAO_VAL[p.c_barra] || 0),
    pontos.map(p => p.c_barra), '#7b1fa2');

  desenharLinhaEvolucao('grafFlexao', labels,
    pontos.map(p => MENCAO_VAL[p.c_flexao] || 0),
    pontos.map(p => p.c_flexao), '#e65100');

  desenharLinhaEvolucao('grafAbdom', labels,
    pontos.map(p => MENCAO_VAL[p.c_abd] || 0),
    pontos.map(p => p.c_abd), '#00838f');

  desenharLinhaEvolucao('grafCorrida', labels,
    pontos.map(p => MENCAO_VAL[p.c_corrida] || 0),
    pontos.map(p => p.c_corrida), '#c62828');
}

function desenharLinhaEvolucao(canvasId, labels, valores, mencoes, cor, titulo) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const W   = canvas.width  = canvas.offsetWidth  || 400;
  const H   = canvas.height = canvas.offsetHeight || 200;
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0,0,W,H);
  ctx.fillStyle = '#fff'; ctx.fillRect(0,0,W,H);

  if (!labels.length) return;

  const pad = { t:20, r:20, b:36, l:46 };
  const gW  = W - pad.l - pad.r;
  const gH  = H - pad.t - pad.b;
  const maxV = 5, minV = 0;

  // Grid horizontal
  for (let v = 1; v <= 5; v++) {
    const y = pad.t + gH - ((v - minV) / (maxV - minV)) * gH;
    ctx.strokeStyle = '#eee'; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(pad.l, y); ctx.lineTo(W - pad.r, y); ctx.stroke();
    // Label eixo Y
    ctx.fillStyle = '#888'; ctx.font = '10px Barlow,sans-serif'; ctx.textAlign = 'right';
    ctx.fillText(MENCAO_LABEL[v], pad.l - 4, y + 4);
  }

  // Eixos
  ctx.strokeStyle = '#bbb'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(pad.l, pad.t); ctx.lineTo(pad.l, H - pad.b); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(pad.l, H - pad.b); ctx.lineTo(W - pad.r, H - pad.b); ctx.stroke();

  const xStep = labels.length > 1 ? gW / (labels.length - 1) : gW / 2;

  const pts = valores.map((v, i) => ({
    x: pad.l + (labels.length > 1 ? i * xStep : gW / 2),
    y: pad.t + gH - ((Math.max(0, v) - minV) / (maxV - minV)) * gH,
    v, m: mencoes[i],
  }));

  // Área preenchida
  if (pts.length > 1) {
    ctx.beginPath();
    ctx.moveTo(pts[0].x, H - pad.b);
    pts.forEach(p => ctx.lineTo(p.x, p.y));
    ctx.lineTo(pts[pts.length-1].x, H - pad.b);
    ctx.closePath();
    ctx.fillStyle = cor + '22'; ctx.fill();
  }

  // Linha
  if (pts.length > 1) {
    ctx.beginPath(); ctx.moveTo(pts[0].x, pts[0].y);
    pts.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
    ctx.strokeStyle = cor; ctx.lineWidth = 2.5; ctx.stroke();
  }

  // Pontos e labels
  pts.forEach((p, i) => {
    // Ponto
    ctx.beginPath();
    ctx.arc(p.x, p.y, 5, 0, Math.PI*2);
    ctx.fillStyle = COR_MENCAO[p.m] || cor; ctx.fill();
    ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.5; ctx.stroke();

    // Menção acima do ponto
    if (p.m) {
      ctx.fillStyle = COR_MENCAO[p.m] || '#333';
      ctx.font = 'bold 11px Barlow,sans-serif'; ctx.textAlign = 'center';
      ctx.fillText(p.m, p.x, p.y - 8);
    }

    // Label eixo X
    ctx.fillStyle = '#555'; ctx.font = '10px Barlow,sans-serif'; ctx.textAlign = 'center';
    ctx.fillText(labels[i], p.x, H - pad.b + 14);
  });
}

// ── Foto ──────────────────────────────────────────────────────────────────
function onSelecionarFoto(event, mid) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    _fotosCache[mid] = e.target.result;
    carregarFicha();
  };
  reader.readAsDataURL(file);
}

function onDropFoto(event, mid) {
  event.preventDefault();
  document.getElementById('fichaFotoBox')?.classList.remove('foto-drag');
  const file = event.dataTransfer.files[0];
  if (!file || !file.type.startsWith('image/')) return;
  const reader = new FileReader();
  reader.onload = (e) => { _fotosCache[mid] = e.target.result; carregarFicha(); };
  reader.readAsDataURL(file);
}

function removerFoto(mid) {
  delete _fotosCache[mid];
  carregarFicha();
}

// ── Impressão longitudinal — VERSÃO CORRIGIDA com SVG inline ─────────────
async function imprimirFicha() {
  const mid = _fichaMilAtual;
  if (!mid) { toast('Selecione um militar primeiro.', true); return; }

  const m = state.militares.find(x => x.id === mid);
  if (!m) return;

  const resultados = await window.api.resultados.porMilitar(mid);
  const ORDEM_TIPO = ['diagnostica','1taf','2taf','3taf'];
  resultados.sort((a,b) => a.ano - b.ano || ORDEM_TIPO.indexOf(a.tipo) - ORDEM_TIPO.indexOf(b.tipo));

  const armaObj = getArma(m.arma);
  const padrao  = calcularPadrao(m.arma, m.situacao_funcional);
  const fotoSrc = _fotosCache[mid] || null;

  // Labels das avaliações no eixo X
  const labels = resultados.map(r =>
    r.tipo === 'diagnostica' ? 'AD' : r.tipo.replace('taf','ºAC')
  );

  // Gerar SVGs antes de abrir a janela
  const svgGlobal = resultados.length > 0
    ? svgLinhaEvolucao(labels, resultados.map(r => r.conceito_global), 800, 160)
    : '';

  const OIIs = [
    { campo:'conceito_barra',         label:'Barra Fixa' },
    { campo:'conceito_flexao_bracos', label:'Flexão de Braço' },
    { campo:'conceito_abdominal',     label:'Abdominal' },
    { campo:'conceito_corrida',       label:'Corrida 12min' },
  ];
  const svgsOII = OIIs.map(o =>
    svgLinhaEvolucao(labels, resultados.map(r => r[o.campo]), 280, 150)
  );

  // Tabela HTML inline (não depende de DOM da janela mãe)
  const tabelaHTML = resultados.map((r,i) => `
    <tr>
      <td>${r.descricao}</td>
      <td>${labelTipo(r.tipo)}</td>
      <td>${r.ano}</td>
      <td>${r.chamada}ª</td>
      <td>${r.corrida_12min||'—'}</td><td>${r.conceito_corrida||''}</td>
      <td>${r.flexao_bracos||'—'}</td><td>${r.conceito_flexao_bracos||''}</td>
      <td>${r.abdominal_supra||'—'}</td><td>${r.conceito_abdominal||''}</td>
      <td>${r.flexao_barra||'—'}</td><td>${r.conceito_barra||''}</td>
      <td><b>${r.conceito_global||'NR'}</b></td>
      <td>${r.suficiencia||'NR'}</td>
    </tr>`).join('');

  const win = window.open('', '_blank');
  if (!win) { toast('Permita popups para imprimir.', true); return; }

  win.document.write(`<!DOCTYPE html><html lang="pt-BR"><head>
  <meta charset="UTF-8">
  <title>Ficha TAF — ${m.posto_graduacao} ${m.nome_guerra||m.nome}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@400;500&display=swap');
    @page { size: A4 landscape; margin: 8mm 10mm; }
    * { margin:0;padding:0;box-sizing:border-box; }
    body { font-family:'Barlow',sans-serif;font-size:10px;color:#1a2010; }

    /* Cabeçalho */
    .topo { display:flex;align-items:flex-start;gap:14px;
            border-bottom:2px solid #c9a227;padding-bottom:8px;margin-bottom:10px; }
    .foto { width:68px;height:85px;object-fit:cover;border:1px solid #bbb;border-radius:3px;flex-shrink:0; }
    .foto-ph { width:68px;height:85px;background:#f0ede6;border:1px dashed #ccc;border-radius:3px;
               flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:26px;color:#ccc; }
    .dados h1 { font-family:'Barlow Condensed',sans-serif;font-size:17px;color:#2d4a1e;letter-spacing:.5px; }
    .dados h2 { font-family:'Barlow Condensed',sans-serif;font-size:10px;color:#888;font-weight:400;margin-bottom:5px; }
    .dados-grid { display:flex;flex-wrap:wrap;gap:4px 16px; }
    .dados-item label { display:block;font-size:7.5px;text-transform:uppercase;letter-spacing:.8px;color:#aaa; }
    .dados-item strong { font-size:9px; }

    /* Gráficos */
    .sec { font-family:'Barlow Condensed',sans-serif;font-size:10px;letter-spacing:.5px;
           color:#2d4a1e;border-bottom:1px solid #e0d8c8;padding-bottom:2px;
           margin:8px 0 5px;text-transform:uppercase; }
    .graf-global svg { width:100%;height:auto;display:block; }
    .oii-grid { display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-bottom:8px; }
    .oii-item label { display:block;text-align:center;font-size:8px;color:#888;margin-bottom:2px; }
    .oii-item svg { width:100%;height:auto;display:block; }

    /* Tabela */
    table { width:100%;border-collapse:collapse;font-size:8.5px; }
    th { background:#2d4a1e;color:#c9a227;padding:3px 5px;
         font-family:'Barlow Condensed',sans-serif;font-size:9px;letter-spacing:.3px; }
    td { padding:2.5px 5px;border-bottom:1px solid #ede7d8; }
    tr:nth-child(even) td { background:#f9f6f0; }
    .portaria { font-size:7px;color:#aaa;margin-top:6px;border-top:1px solid #e0d8c8;
                padding-top:3px;text-align:center; }
  </style>
  </head><body>

  <div class="topo">
    ${fotoSrc ? `<img class="foto" src="${fotoSrc}">` : '<div class="foto-ph">👤</div>'}
    <div class="dados">
      <h1>${m.posto_graduacao} ${armaObj?.sigla||''} ${m.nome_guerra||m.nome} — FICHA INDIVIDUAL DE DESEMPENHO FÍSICO</h1>
      <h2>${m.nome}</h2>
      <div class="dados-grid">
        <div class="dados-item"><label>Nascimento</label><strong>${formatarData(m.data_nascimento)}</strong></div>
        <div class="dados-item"><label>Idade</label><strong>${TAF.calcularIdade(m.data_nascimento)} anos</strong></div>
        <div class="dados-item"><label>Sexo</label><strong>${m.sexo==='M'?'Masculino':'Feminino'}</strong></div>
        <div class="dados-item"><label>Arma</label><strong>${armaObj?.sigla||m.arma||'—'} — ${armaObj?.label||''}</strong></div>
        <div class="dados-item"><label>LEM</label><strong>${m.lem||'—'}</strong></div>
        <div class="dados-item"><label>Padrão</label><strong>${padrao}</strong></div>
        <div class="dados-item"><label>OM</label><strong>${m.om||'—'}</strong></div>
        <div class="dados-item"><label>Subunidade</label><strong>${m.subunidade||'—'}</strong></div>
      </div>
    </div>
  </div>

  ${resultados.length === 0 ? '<p style="color:#888;text-align:center;padding:20px">Nenhum resultado registrado.</p>' : `
  <div class="sec">Evolução — Menção Global (AD → 1º AC)</div>
  <div class="graf-global">${svgGlobal}</div>

  <div class="sec">Evolução por OII</div>
  <div class="oii-grid">
    ${OIIs.map((o,i) => `
    <div class="oii-item">
      <label>${o.label}</label>
      ${svgsOII[i]}
    </div>`).join('')}
  </div>

  <div class="sec">Histórico de Avaliações</div>
  <table>
    <thead><tr>
      <th>Avaliação</th><th>Tipo</th><th>Ano</th><th>Cham.</th>
      <th>Corrida</th><th>M</th><th>Flexão</th><th>M</th>
      <th>Abdom</th><th>M</th><th>Barra</th><th>M</th>
      <th>Global</th><th>Sufic.</th>
    </tr></thead>
    <tbody>${tabelaHTML}</tbody>
  </table>`}

  <div class="portaria">Portaria EME/C Ex Nº 850, de 31 de agosto de 2022 (EB20-D-03.053) · TAF-EB Sistema de Avaliação Física</div>
  <script>window.onload = function(){ setTimeout(function(){ window.print(); }, 400); };</script>
  </body></html>`);
  win.document.close();
}


// ═════════════════════════════════════════════════════════════════════════
// IMPORTAR CSV — detecção automática: Modo A = Militares / Modo B = Resultados
// ═════════════════════════════════════════════════════════════════════════

let csvPreviewData    = [];   // resultados parseados (Modo B)
let _csvMilitaresData = [];   // militares parseados  (Modo A)
let _csvModo          = null; // 'militares' | 'resultados'

function paginaImportar() {
  const armasLista = ARMAS.map(a => a.value).join(' · ');
  return `
  <div class="page-header">
    <div>
      <div class="page-title">IMPORTAR CSV</div>
      <div class="page-subtitle">Modo detectado automaticamente pelo cabeçalho do arquivo</div>
    </div>
    <div class="page-actions">
      <button class="btn btn-ghost" onclick="baixarModeloCSV('militares')">⬇ Modelo Militares</button>
      <button class="btn btn-ghost" onclick="baixarModeloCSV('resultados')">⬇ Modelo Resultados</button>
    </div>
  </div>

  <div class="card">
    <div class="card-title">📋 FORMATOS ACEITOS</div>

    <div class="secao-titulo">MODO A — MILITARES (cadastra no sistema)</div>
    <code style="display:block;background:#f4f4f4;padding:8px 12px;border-radius:4px;font-size:11px;overflow-x:auto;line-height:2;margin-bottom:6px">
      posto;om;nome;nome_guerra;data_nascimento;sexo;arma;situacao_funcional;subunidade<br>
      Cap;CAV;FABRICIO BITTENCOURT MORAES;MORAES;02/02/1991;M;CAV;OM_Op;
    </code>
    <div style="font-size:11px;color:#666;margin-bottom:12px">
      <strong>arma:</strong> ${armasLista} &nbsp;|&nbsp;
      <strong>situacao_funcional:</strong> OM_NOp · OM_Op · OM_F_Emp_Estrt · Estb_Ens &nbsp;|&nbsp;
      <strong>nome_guerra, om, subunidade:</strong> opcionais
    </div>

    <div class="secao-titulo">MODO B — RESULTADOS (vincula a militares já cadastrados)</div>
    <code style="display:block;background:#f4f4f4;padding:8px 12px;border-radius:4px;font-size:11px;overflow-x:auto;line-height:2;margin-bottom:6px">
      nome;data_nascimento;chamada;corrida;flexao_bracos;abdominal_supra;flexao_barra;ppm_tempo<br>
      MORAES;02/02/1991;1;2900;31;70;9;
    </code>
    <div style="font-size:11px;color:#666">
      <strong>nome:</strong> nome de guerra ou completo &nbsp;|&nbsp;
      <strong>data_nascimento:</strong> opcional, recomendado para homônimos &nbsp;|&nbsp;
      <strong>ppm_tempo:</strong> MM:SS
    </div>
  </div>

  <div class="card">
    <div class="card-title">📂 CARREGAR ARQUIVO</div>
    <div class="form-grid" style="margin-bottom:14px">
      <div class="form-group">
        <label>Avaliação TAF de destino <span style="color:#999">(Modo B)</span></label>
        <select id="imp_aval">
          <option value="">— Selecione se for importar resultados —</option>
          ${state.avaliacoes.map(a => '<option value="' + a.id + '">' + a.descricao + ' (' + a.ano + ')</option>').join('')}
        </select>
      </div>
      <div class="form-group">
        <label>Chamada padrão <span style="color:#999">(Modo B)</span></label>
        <select id="imp_chamada">
          <option value="1">1ª Chamada</option>
          <option value="2">2ª Chamada</option>
          <option value="3">3ª Chamada</option>
          <option value="4">Chamada Extra</option>
        </select>
      </div>
      <div class="form-group">
        <label>Militar já existente <span style="color:#999">(Modo A)</span></label>
        <select id="imp_duplicado">
          <option value="skip">Ignorar (manter cadastro atual)</option>
          <option value="update">Atualizar cadastro</option>
        </select>
      </div>
    </div>
    <div id="dropZone" class="drop-zone"
         ondragover="event.preventDefault();this.classList.add('drag-over')"
         ondragleave="this.classList.remove('drag-over')"
         ondrop="onDropCSV(event)">
      <div class="drop-icon">⬆</div>
      <div class="drop-texto">Arraste o arquivo CSV aqui</div>
      <div class="drop-sub">ou</div>
      <label class="btn btn-primary" style="cursor:pointer;margin-top:6px">
        Selecionar arquivo
        <input type="file" accept=".csv,.txt" style="display:none" onchange="onFileCSV(event)">
      </label>
    </div>
  </div>

  <div id="previewArea"></div>
  <div id="btnConfirmar" style="display:none;text-align:right;margin-top:12px">
    <button class="btn btn-ghost" onclick="limparImport()">✕ Cancelar</button>
    <button class="btn btn-ouro" id="btnConfirmarLabel" onclick="confirmarImport()">✔ Confirmar</button>
  </div>`;
}

function onDropCSV(event) {
  event.preventDefault();
  document.getElementById('dropZone').classList.remove('drag-over');
  const file = event.dataTransfer.files[0];
  if (file) lerArquivoCSV(file);
}
function onFileCSV(event) {
  const file = event.target.files[0];
  if (file) lerArquivoCSV(file);
}
function lerArquivoCSV(file) {
  const reader = new FileReader();
  reader.onload = (e) => processarCSV(e.target.result, file.name);
  reader.readAsText(file, 'UTF-8');
}

// ── Normalização e detecção de modo ────────────────────────────────────
const normStr = s => (s || '').toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').trim();

function detectarModo(cabecalho) {
  const colsMil = ['posto', 'arma', 'situacao_funcional', 'posto_graduacao', 'tipo_om'];
  return colsMil.some(c => cabecalho.includes(c)) ? 'militares' : 'resultados';
}

function parsearLinhas(texto) {
  return texto.split('\n')
    .map(l => l.replace(/\r/g, '').trim())
    .filter(l => l && !l.startsWith('#'));
}

function detectarSep(linha) {
  return (linha.match(/;/g)||[]).length >= (linha.match(/,/g)||[]).length ? ';' : ',';
}

function mapIdx(cabecalho, MAP) {
  const idx = {};
  for (const [campo, aliases] of Object.entries(MAP)) {
    idx[campo] = -1;
    for (const a of aliases) {
      const i = cabecalho.indexOf(a);
      if (i >= 0) { idx[campo] = i; break; }
    }
  }
  return idx;
}

// ── Parser Modo A: Militares ────────────────────────────────────────────
function parseMilitares(linhas, sep) {
  const cab = linhas[0].split(sep).map(c => normStr(c));
  const MAP = {
    posto:              ['posto','posto_graduacao','grad','graduacao'],
    om:                 ['om','organizacao_militar','unidade'],
    nome:               ['nome','nome_completo'],
    nome_guerra:        ['nome_guerra','guerra','nome_de_guerra'],
    data_nascimento:    ['data_nascimento','nascimento','data_nasc','dt_nasc','dn'],
    sexo:               ['sexo','sex'],
    arma:               ['arma','arma_quadro','quadro','servico'],
    situacao_funcional: ['situacao_funcional','sit_funcional','situacao','om_tipo','tipo_om'],
    subunidade:         ['subunidade','su'],
  };
  const idx = mapIdx(cab, MAP);
  if (idx.nome < 0) return { erro: 'Coluna "nome" não encontrada. Colunas lidas: [' + cab.join(' | ') + ']' };

  const SITS = ['OM_NOp','OM_Op','OM_F_Emp_Estrt','Estb_Ens'];
  const erros = [], validos = [];

  for (let i = 1; i < linhas.length; i++) {
    const cols = linhas[i].split(sep).map(c => c.trim());
    const lNum = i + 1;
    const get  = f => idx[f] >= 0 && idx[f] < cols.length ? cols[idx[f]].trim() : '';

    const nome = get('nome');
    if (!nome) { erros.push('Linha ' + lNum + ': nome vazio.'); continue; }

    let dataNasc = get('data_nascimento');
    if (dataNasc.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
      const [d,m,y] = dataNasc.split('/'); dataNasc = y+'-'+m+'-'+d;
    } else if (dataNasc && !dataNasc.match(/^\d{4}-\d{2}-\d{2}$/)) {
      erros.push('Linha ' + lNum + ': data_nascimento inválida ("' + dataNasc + '"). Use DD/MM/AAAA.'); continue;
    }
    if (!dataNasc) { erros.push('Linha ' + lNum + ': data_nascimento obrigatória.'); continue; }

    const sexo = get('sexo').toUpperCase()
      .replace('MASCULINO','M').replace('FEMININO','F').replace('MASC','M').replace('FEM','F');
    if (!['M','F'].includes(sexo)) { erros.push('Linha ' + lNum + ': sexo inválido ("' + get('sexo') + '"). Use M ou F.'); continue; }

    const armaRaw = get('arma').toUpperCase();
    const armaObj = getArma(armaRaw);
    if (!armaObj) {
      erros.push('Linha ' + lNum + ': arma "' + armaRaw + '" inválida. Valores: ' + ARMAS.map(a=>a.value).join(', ')); continue;
    }

    const sitRaw = get('situacao_funcional');
    if (sitRaw && !SITS.includes(sitRaw)) {
      erros.push('Linha ' + lNum + ': situacao_funcional "' + sitRaw + '" inválida. Use: ' + SITS.join(' · ')); continue;
    }

    validos.push({ lNum, militar: {
      nome, nome_guerra: get('nome_guerra')||null,
      posto_graduacao: get('posto')||'Sd',
      arma: armaObj.value, lem: armaObj.lem,
      situacao_funcional: sitRaw||'OM_NOp',
      sexo, data_nascimento: dataNasc,
      om: get('om')||null, subunidade: get('subunidade')||null,
    }});
  }
  return { modo: 'militares', validos, erros };
}

// ── Parser Modo B: Resultados ───────────────────────────────────────────
function parseResultados(linhas, sep) {
  const cab = linhas[0].split(sep).map(c => normStr(c));
  const MAP = {
    nome:            ['nome','nome_guerra','guerra','nome_completo','militar'],
    data_nascimento: ['data_nascimento','nascimento','data_nasc','dt_nasc','dn'],
    chamada:         ['chamada','n_chamada','num_chamada'],
    corrida:         ['corrida','corrida_12min','corrida_12','corrida12'],
    flexao_bracos:   ['flexao_bracos','flexao','flex_bracos','flexao_de_bracos'],
    abdominal_supra: ['abdominal_supra','abdominal','abd','abdomen'],
    flexao_barra:    ['flexao_barra','barra','barra_fixa','flex_barra'],
    ppm_tempo:       ['ppm_tempo','ppm','pista','tempo_ppm'],
  };
  const idx = mapIdx(cab, MAP);
  if (idx.nome < 0) return { erro: 'Coluna "nome" não encontrada. Colunas lidas: [' + cab.join(' | ') + ']' };
  if (!state.militares.length) return { erro: 'Nenhum militar cadastrado. Importe os militares primeiro (Modo A).' };

  const chamadaPad = parseInt(document.getElementById('imp_chamada')?.value) || 1;
  const erros = [], validos = [];
  const milNorm = state.militares.map(m => ({
    m, guerra: normStr(m.nome_guerra||''), compl: normStr(m.nome),
  }));

  for (let i = 1; i < linhas.length; i++) {
    const cols = linhas[i].split(sep).map(c => c.trim());
    const lNum = i + 1;
    const get  = f => idx[f] >= 0 && idx[f] < cols.length ? cols[idx[f]].trim() : '';

    const nomeCSV = get('nome');
    if (!nomeCSV) { erros.push('Linha ' + lNum + ': nome vazio.'); continue; }

    let dataNasc = get('data_nascimento');
    if (dataNasc.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
      const [d,m,y] = dataNasc.split('/'); dataNasc = y+'-'+m+'-'+d;
    } else if (dataNasc && !dataNasc.match(/^\d{4}-\d{2}-\d{2}$/)) {
      erros.push('Linha ' + lNum + ': data_nascimento inválida ("' + dataNasc + '").'); continue;
    }

    const nN = normStr(nomeCSV);
    const militar =
      milNorm.find(x => x.guerra === nN && (!dataNasc || x.m.data_nascimento === dataNasc))?.m ||
      milNorm.find(x => x.compl  === nN && (!dataNasc || x.m.data_nascimento === dataNasc))?.m ||
      (!dataNasc && milNorm.find(x => x.guerra === nN)?.m) ||
      (!dataNasc && milNorm.find(x => x.compl  === nN)?.m) ||
      (dataNasc  && milNorm.find(x => (x.guerra.includes(nN)||x.compl.includes(nN)) && x.m.data_nascimento === dataNasc)?.m);

    if (!militar) {
      erros.push('Linha ' + lNum + ': "' + nomeCSV + '"' + (dataNasc ? ' (nasc. ' + dataNasc + ')' : '') + ' — não encontrado no cadastro.');
      continue;
    }

    const armaObj = getArma(militar.arma);
    const lem     = armaObj?.lem || militar.lem || 'LEMB';

    const resultado = {
      militar_id:      militar.id,
      chamada:         parseInt(get('chamada')) || chamadaPad,
      corrida_12min:   get('corrida')          ? parseInt(get('corrida'))          : null,
      flexao_bracos:   get('flexao_bracos')    ? parseInt(get('flexao_bracos'))    : null,
      abdominal_supra: get('abdominal_supra')  ? parseInt(get('abdominal_supra'))  : null,
      flexao_barra:    get('flexao_barra')     ? parseInt(get('flexao_barra'))     : null,
      ppm_tempo:       get('ppm_tempo')        || null,
      _militar: militar, _lem: lem,
    };
    const ppmSeg = TAF.mmssParaSegundos(resultado.ppm_tempo);
    resultado._calc = TAF.calcularTAF({
      dataNascimento: militar.data_nascimento, lem,
      sexo: militar.sexo, situacaoFuncional: militar.situacao_funcional,
      postoGraduacao: militar.posto_graduacao,
      corrida: resultado.corrida_12min, flexao: resultado.flexao_bracos,
      abdominal: resultado.abdominal_supra, barra: resultado.flexao_barra,
      ppmTempo: ppmSeg, tafAlternativo: false,
      // Usar data da chamada da avaliação selecionada para cálculo correto de idade
      dataChamada: (() => {
        const _av = state.avaliacoes.find(a => a.id === parseInt(document.getElementById('imp_aval')?.value));
        const _ch = resultado.chamada || 1;
        return _av ? (_ch===1?_av.data_1_chamada:_ch===2?_av.data_2_chamada:_ch===3?_av.data_3_chamada:_av.data_chamada_extra) : null;
      })(),
      anoTAF: (() => { const _av2=state.avaliacoes.find(a=>a.id===parseInt(document.getElementById('imp_aval')?.value)); return _av2?.ano||null; })(),
      padrao: state.avaliacoes.find(a => a.id === parseInt(document.getElementById('imp_aval')?.value))?.padrao || null,
    });
    validos.push({ lNum, resultado });
  }
  return { modo: 'resultados', validos, erros };
}

// ── Dispatch ────────────────────────────────────────────────────────────
function parseCSV(texto) {
  const linhas = parsearLinhas(texto);
  if (linhas.length < 2) return { erro: 'Arquivo vazio ou sem dados. Verifique o cabeçalho.' };
  const sep = detectarSep(linhas[0]);
  const cab = linhas[0].split(sep).map(c => normStr(c));
  const modo = detectarModo(cab);
  return modo === 'militares' ? parseMilitares(linhas, sep) : parseResultados(linhas, sep);
}

// ── Prévia ──────────────────────────────────────────────────────────────
function processarCSV(texto, nomeArquivo) {
  const parsed = parseCSV(texto);
  csvPreviewData = []; _csvMilitaresData = []; _csvModo = parsed.modo || null;

  if (parsed.erro) {
    document.getElementById('previewArea').innerHTML =
      '<div class="card" style="border-left:3px solid #e74c3c"><p style="color:#c0392b;font-size:13px">❌ ' + parsed.erro + '</p></div>';
    return;
  }

  const errosHTML = parsed.erros.length ? `
    <div class="card" style="border-left:3px solid #e74c3c;margin-bottom:12px">
      <div class="card-title" style="color:#c0392b">⚠ ${parsed.erros.length} linha(s) com problema</div>
      <ul style="font-size:12px;padding-left:18px;color:#721c24;max-height:160px;overflow-y:auto">
        ${parsed.erros.map(e => '<li>' + e + '</li>').join('')}
      </ul>
    </div>` : '';

  if (!parsed.validos.length) {
    document.getElementById('previewArea').innerHTML = errosHTML +
      '<div class="card"><p style="color:#856404">⚠ Nenhuma linha válida encontrada.</p></div>';
    document.getElementById('btnConfirmar').style.display = 'none';
    return;
  }

  let tabelaHTML = '';
  const modoLabel = parsed.modo === 'militares' ? '👤 MODO A — MILITARES' : '📊 MODO B — RESULTADOS';

  if (parsed.modo === 'militares') {
    _csvMilitaresData = parsed.validos;
    const cores = { PBD:'badge-B', PAD:'badge-MB', PED:'badge-E' };
    tabelaHTML = '<table><thead><tr>' +
      '<th>#</th><th>Posto</th><th>OM</th><th>Nome</th><th>Guerra</th>' +
      '<th class="center">Nasc.</th><th class="center">Sx</th><th class="center">Idade</th>' +
      '<th>Arma</th><th class="center">LEM</th><th class="center">Padrão</th><th>Tipo OM</th><th class="center">Status</th>' +
      '</tr></thead><tbody>' +
      parsed.validos.slice(0,100).map(({lNum, militar: m}) => {
        const ao = getArma(m.arma);
        const pad = calcularPadrao(m.arma, m.situacao_funcional);
        const existe = state.militares.find(x => x.nome.toLowerCase()===m.nome.toLowerCase() && x.data_nascimento===m.data_nascimento);
        return '<tr>' +
          '<td class="text-muted">' + lNum + '</td>' +
          '<td>' + m.posto_graduacao + '</td>' +
          '<td>' + (m.om||'—') + '</td>' +
          '<td><strong>' + m.nome + '</strong></td>' +
          '<td>' + (m.nome_guerra||'—') + '</td>' +
          '<td class="center">' + formatarData(m.data_nascimento) + '</td>' +
          '<td class="center">' + m.sexo + '</td>' +
          '<td class="center">' + TAF.calcularIdade(m.data_nascimento) + '</td>' +
          '<td><span class="badge badge-B" style="font-size:10px">' + (ao?.sigla||m.arma) + '</span> ' + (ao?.label||'') + '</td>' +
          '<td class="center"><span class="badge badge-B" style="font-size:10px">' + m.lem + '</span></td>' +
          '<td class="center"><span class="badge ' + (cores[pad]||'badge-B') + '" style="font-size:10px">' + pad + '</span></td>' +
          '<td>' + labelSitFuncional(m.situacao_funcional) + '</td>' +
          '<td class="center">' + (existe
            ? '<span style="color:#856404;font-size:11px">✎ atualizar</span>'
            : '<span style="color:#155724;font-size:11px">✚ novo</span>') + '</td>' +
          '</tr>';
      }).join('') +
      '</tbody></table>';
    document.getElementById('btnConfirmarLabel').textContent = '✔ Cadastrar ' + parsed.validos.length + ' Militar(es)';

  } else {
    csvPreviewData = parsed.validos;
    const cPad = { PBD:'badge-B', PAD:'badge-MB', PED:'badge-E' };
    const bc = c => c ? '<span class="badge badge-' + c + '" style="font-size:10px">' + c + '</span>' : '';
    tabelaHTML = '<table><thead><tr>' +
      '<th>#</th><th>Posto</th><th>Nome</th>' +
      '<th class="center">Id</th><th class="center">LEM</th><th class="center">Pad</th>' +
      '<th class="center">Cham.</th><th class="center">Corrida</th><th class="center">Flexão</th>' +
      '<th class="center">Abdom</th><th class="center">Barra</th><th class="center">PPM</th>' +
      '<th class="center">Global</th><th class="center">Sufic.</th>' +
      '</tr></thead><tbody>' +
      parsed.validos.slice(0,100).map(({lNum, resultado: r}) => {
        const m = r._militar, c = r._calc;
        return '<tr>' +
          '<td class="text-muted">' + lNum + '</td>' +
          '<td>' + m.posto_graduacao + '</td>' +
          '<td><strong>' + (m.nome_guerra||m.nome) + '</strong></td>' +
          '<td class="center">' + TAF.calcularIdade(m.data_nascimento) + '</td>' +
          '<td class="center"><span class="badge badge-B" style="font-size:10px">' + r._lem + '</span></td>' +
          '<td class="center"><span class="badge ' + (cPad[c.padrao]||'badge-B') + '" style="font-size:10px">' + c.padrao + '</span></td>' +
          '<td class="center">' + r.chamada + 'ª</td>' +
          '<td class="center">' + (r.corrida_12min??'—') + ' ' + bc(c.conceitos?.corrida) + '</td>' +
          '<td class="center">' + (r.flexao_bracos??'—') + ' ' + bc(c.conceitos?.flexao) + '</td>' +
          '<td class="center">' + (r.abdominal_supra??'—') + ' ' + bc(c.conceitos?.abdominal) + '</td>' +
          '<td class="center">' + (r.flexao_barra??'—') + ' ' + bc(c.conceitos?.barra) + '</td>' +
          '<td class="center">' + (r.ppm_tempo||'—') + ' ' + bc(c.suficiencias?.ppm) + '</td>' +
          '<td class="center"><strong>' + badgeConceito(c.conceitoGlobal) + '</strong></td>' +
          '<td class="center">' + badgeSuficiencia(c.suficienciaGlobal) + '</td>' +
          '</tr>';
      }).join('') +
      '</tbody></table>';
    document.getElementById('btnConfirmarLabel').textContent = '✔ Importar ' + parsed.validos.length + ' Resultado(s)';
  }

  document.getElementById('previewArea').innerHTML = errosHTML + `
    <div class="card" style="padding:0;overflow:hidden">
      <div style="padding:10px 16px;background:var(--verde-eb);display:flex;align-items:center;gap:12px;flex-wrap:wrap">
        <span style="color:var(--ouro-claro);font-family:'Barlow Condensed',sans-serif;font-size:16px;letter-spacing:.5px">
          ${modoLabel} — ${nomeArquivo} · ${parsed.validos.length} válida(s)
          ${parsed.erros.length ? '<span style="color:#e87070;font-size:12px;margin-left:8px">(' + parsed.erros.length + ' com problema)</span>' : ''}
        </span>
      </div>
      <div class="table-wrap">${tabelaHTML}</div>
    </div>`;

  document.getElementById('btnConfirmar').style.display = 'block';
}

// ── Confirmar importação ─────────────────────────────────────────────────
async function confirmarImport() {
  if (_csvModo === 'militares') {
    if (!_csvMilitaresData.length) { toast('Nenhum dado para importar.', true); return; }
    const dup = document.getElementById('imp_duplicado')?.value || 'skip';
    let novos = 0, atu = 0;
    for (const { militar } of _csvMilitaresData) {
      const ex = state.militares.find(x =>
        x.nome.toLowerCase() === militar.nome.toLowerCase() && x.data_nascimento === militar.data_nascimento);
      if (ex) {
        if (dup === 'update') { await window.api.militares.salvar({...militar, id: ex.id}); atu++; }
      } else {
        await window.api.militares.salvar(militar); novos++;
      }
    }
    limparImport(); await carregarDados();
    toast('✔ ' + novos + ' militar(es) cadastrado(s)' + (atu ? ' · ' + atu + ' atualizado(s)' : '') + '.');

  } else {
    if (!csvPreviewData.length) { toast('Nenhum dado para importar.', true); return; }
    const avalId = parseInt(document.getElementById('imp_aval')?.value);
    if (!avalId) {
      toast('Selecione a avaliação TAF de destino.', true);
      document.getElementById('imp_aval').style.border = '2px solid red';
      document.getElementById('imp_aval').focus(); return;
    }
    let salvos = 0;
    for (const { resultado: r } of csvPreviewData) {
      const c = r._calc;
      await window.api.resultados.salvar({
        militar_id: r.militar_id, avaliacao_id: avalId, chamada: r.chamada,
        corrida_12min: r.corrida_12min, flexao_bracos: r.flexao_bracos,
        abdominal_supra: r.abdominal_supra, flexao_barra: r.flexao_barra, ppm_tempo: r.ppm_tempo,
        conceito_corrida: c.conceitos?.corrida||null, conceito_flexao_bracos: c.conceitos?.flexao||null,
        conceito_abdominal: c.conceitos?.abdominal||null, conceito_barra: c.conceitos?.barra||null,
        conceito_ppm: c.suficiencias?.ppm||null, conceito_global: c.conceitoGlobal,
        suficiencia: c.suficienciaGlobal, padrao_verificado: c.padrao,
        taf_alternativo: 0, observacoes: null,
      });
      salvos++;
    }
    limparImport(); await carregarDados();
    toast('✔ ' + salvos + ' resultado(s) importados com sucesso!');
  }
}

function limparImport() {
  csvPreviewData = []; _csvMilitaresData = []; _csvModo = null;
  document.getElementById('previewArea').innerHTML = '';
  document.getElementById('btnConfirmar').style.display = 'none';
}

function baixarModeloCSV(tipo) {
  let linhas, nome;
  if (tipo === 'militares') {
    linhas = [
      '# TAF-EB — Modelo MILITARES (Modo A)',
      '# arma: INF | CAV | ART | ENG | COM | AvEx | TOP | SSau | MatBel | Int | QEM | QAO | QCO | SAREX | MUS | LEMCT',
      '# situacao_funcional: OM_NOp | OM_Op | OM_F_Emp_Estrt | Estb_Ens',
      '# sexo: M ou F | data_nascimento: DD/MM/AAAA | nome_guerra, om, subunidade: opcionais',
      'posto;om;nome;nome_guerra;data_nascimento;sexo;arma;situacao_funcional;subunidade',
      'Cap;CAV;FABRICIO BITTENCOURT MORAES;MORAES;02/02/1991;M;CAV;OM_Op;',
      'Cap;CAV;FRANCISCO MELLO SIQUEIRA NETO;MELLO;03/06/1991;M;CAV;OM_Op;',
      'Cap;CAV;GUILHERME ALBERTI BRESSAN;BRESSAN;31/10/1994;M;CAV;OM_Op;',
      'Cap;SSau;MARIA SILVA PEREIRA;SILVA;15/04/1993;F;SSau;OM_NOp;',
    ];
    nome = 'modelo-militares-taf-eb.csv';
  } else {
    linhas = [
      '# TAF-EB — Modelo RESULTADOS (Modo B)',
      '# nome: nome de guerra ou completo | ppm_tempo: MM:SS | chamada: 1,2,3 ou 4',
      'nome;data_nascimento;chamada;corrida;flexao_bracos;abdominal_supra;flexao_barra;ppm_tempo',
      'MORAES;02/02/1991;1;2900;31;70;9;',
      'MELLO;03/06/1991;1;3000;40;70;10;',
      'BRESSAN;31/10/1994;1;2750;31;80;;',
      'SILVA;15/04/1993;1;2300;14;55;;',
    ];
    nome = 'modelo-resultados-taf-eb.csv';
  }
  const blob = new Blob([linhas.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = nome; a.click();
  URL.revokeObjectURL(url);
  toast('Modelo ' + (tipo==='militares'?'Militares':'Resultados') + ' baixado!');
}

// ═════════════════════════════════════════════════════════════════════════
// UTILITÁRIOS DE UI
// ═════════════════════════════════════════════════════════════════════════
function badgeConceito(c) {
  if (!c) return '';
  return `<span class="badge badge-${c}">${c}</span>`;
}

function badgeSuficiencia(s) {
  if (!s) return '—';
  return `<span class="badge badge-${s}">${s}</span>`;
}

function labelTipo(tipo) {
  return TIPOS_TAF.find(t => t.value === tipo)?.label || tipo;
}

function labelSitFuncional(sf) {
  return SITUACOES_FUNCIONAIS.find(s => s.value === sf)?.label || sf;
}

function formatarData(d) {
  if (!d) return '—';
  try {
    const [y, m, dia] = d.split('-');
    return `${dia}/${m}/${y}`;
  } catch { return d; }
}

// ── Modal ─────────────────────────────────────────────────────────────────
function abrirModal(titulo, corpo, rodape) {
  document.getElementById('modalTitulo').textContent = titulo;
  document.getElementById('modalCorpo').innerHTML = corpo;
  document.getElementById('modalRodape').innerHTML = rodape || '';
  document.getElementById('modalOverlay').classList.add('aberto');
  // Focar primeiro input/select após renderizar
  setTimeout(() => {
    const el = document.querySelector('#modalCorpo input:not([type=hidden]):not([type=date]), #modalCorpo select');
    if (el) el.focus();
  }, 60);
}

function fecharModal(event) {
  // Só fecha se clicou diretamente no overlay (fundo escuro), não dentro do modal-box
  if (event && event.target !== document.getElementById('modalOverlay')) return;
  document.getElementById('modalOverlay').classList.remove('aberto');
}

// Fechar modal pelo botão X (sem event)
function fecharModalForce() {
  document.getElementById('modalOverlay').classList.remove('aberto');
}

// ── Toast ─────────────────────────────────────────────────────────────────
function toast(msg, erro) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = 'toast visivel' + (erro ? ' erro' : '');
  setTimeout(() => el.classList.remove('visivel'), 3000);
}

// ── Backup ────────────────────────────────────────────────────────────────
async function backupDb() {
  const res = await window.api.db.backup();
  if (res.ok) toast('Backup salvo em: ' + res.path);
  else toast('Backup cancelado.', true);
}

// ── Navegação com contexto ────────────────────────────────────────────────
function navegarComAval(pagina, avalId) {
  if (pagina === 'lancar') {
    _tafAvalPendente = avalId;  // será lido por renderizarPagina
  }
  navegarPara(pagina);
}

function navegarComMilitar(pagina, milId) {
  navegarPara(pagina);
  setTimeout(() => { carregarFichaMilitar(milId); }, 150);
}
