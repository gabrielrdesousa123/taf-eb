const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

// ── Corrigir erros de cache de GPU (Electron no Windows) ──────────────────
// "Unable to move the cache: Acesso negado" — desabilita cache de GPU
// que tenta escrever em pasta protegida pelo sistema
app.commandLine.appendSwitch('disable-gpu-shader-disk-cache');
app.commandLine.appendSwitch('disable-software-rasterizer');
app.commandLine.appendSwitch('no-sandbox');

let mainWindow;
let db;

// ── Caminho do banco de dados ──────────────────────────────────────────────
// Modo portátil: banco fica em "dados/" ao lado do executável
// Modo dev: fica em "dados/" ao lado do main.js
function getDbPath() {
  const isPortable = process.env.PORTABLE_EXECUTABLE_DIR; // definido pelo electron-builder portátil
  const baseDir = isPortable
    ? path.join(process.env.PORTABLE_EXECUTABLE_DIR, 'dados')
    : path.join(path.dirname(app.getPath('exe')), 'dados');

  // Em desenvolvimento, usar pasta local
  const devDir = path.join(__dirname, 'dados');
  const dir = fs.existsSync(devDir) || !isPortable
    ? devDir
    : baseDir;

  fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, 'taf-eb.db');
}

// ── Inicializar SQLite via sql.js ──────────────────────────────────────────
async function initDatabase() {
  const initSqlJs = require('sql.js');
  const SQL = await initSqlJs();
  const dbPath = getDbPath();

  if (fs.existsSync(dbPath)) {
    const fileBuffer = fs.readFileSync(dbPath);
    db = new SQL.Database(fileBuffer);
  } else {
    db = new SQL.Database();
  }

  criarTabelas();
  migrarTabelas();
  salvarDb();
}

function salvarDb() {
  const data = db.export();
  const buffer = Buffer.from(data);
  fs.writeFileSync(getDbPath(), buffer);
}

// ── Migração: adiciona colunas novas em bancos existentes ─────────────────
// SQLite não suporta ADD COLUMN IF NOT EXISTS — verificamos via PRAGMA
function migrarTabelas() {
  const colsMilitares = db.exec("PRAGMA table_info(militares)")[0];
  if (!colsMilitares) return; // tabela não existe ainda

  const nomes = colsMilitares.values.map(r => r[1]); // coluna 1 = nome

  // Adicionar coluna 'arma' se não existir
  if (!nomes.includes('arma')) {
    db.run("ALTER TABLE militares ADD COLUMN arma TEXT");
    console.log('[migração] coluna arma adicionada em militares');
  }
  // Adicionar coluna 'subunidade' se não existir
  if (!nomes.includes('subunidade')) {
    db.run("ALTER TABLE militares ADD COLUMN subunidade TEXT");
    console.log('[migração] coluna subunidade adicionada em militares');
  }
  // Garante coluna lem não-null (pode ter sido criada sem default em versões antigas)
  if (!nomes.includes('lem')) {
    db.run("ALTER TABLE militares ADD COLUMN lem TEXT NOT NULL DEFAULT 'LEMB'");
    console.log('[migração] coluna lem adicionada em militares');
  }

  // Migrar tabela avaliacoes — padrao e situacao_funcional
  const colsAval = db.exec("PRAGMA table_info(avaliacoes)")[0];
  if (colsAval) {
    const nomesAval = colsAval.values.map(r => r[1]);
    if (!nomesAval.includes('padrao')) {
      db.run("ALTER TABLE avaliacoes ADD COLUMN padrao TEXT DEFAULT 'PAD'");
      console.log('[migração] coluna padrao adicionada em avaliacoes');
    }
    if (!nomesAval.includes('situacao_funcional')) {
      db.run("ALTER TABLE avaliacoes ADD COLUMN situacao_funcional TEXT DEFAULT 'OM_Op'");
      console.log('[migração] coluna situacao_funcional adicionada em avaliacoes');
    }
  }
  // Garantir que todos os militares têm ativo = 1 (corrige bancos antigos com NULL)
  db.run("UPDATE militares SET ativo = 1 WHERE ativo IS NULL");
  // Garantir que todas as avaliações têm ativa = 1 por padrão (exceto as explicitamente inativas)
  db.run("UPDATE avaliacoes SET ativa = 1 WHERE ativa IS NULL");
}

// ── DDL: Criar tabelas ─────────────────────────────────────────────────────
function criarTabelas() {
  db.run(`
    CREATE TABLE IF NOT EXISTS militares (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT NOT NULL,
      nome_guerra TEXT,
      posto_graduacao TEXT NOT NULL,
      arma TEXT,
      lem TEXT NOT NULL,
      situacao_funcional TEXT NOT NULL,
      sexo TEXT NOT NULL,
      data_nascimento TEXT NOT NULL,
      om TEXT,
      subunidade TEXT,
      ativo INTEGER DEFAULT 1,
      criado_em TEXT DEFAULT (datetime('now','localtime'))
    );

    CREATE TABLE IF NOT EXISTS avaliacoes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      descricao TEXT NOT NULL,
      tipo TEXT NOT NULL,
      ano INTEGER NOT NULL,
      padrao TEXT DEFAULT 'PAD',
      situacao_funcional TEXT DEFAULT 'OM_Op',
      data_1_chamada TEXT,
      data_2_chamada TEXT,
      data_3_chamada TEXT,
      data_chamada_extra TEXT,
      ativa INTEGER DEFAULT 1,
      criado_em TEXT DEFAULT (datetime('now','localtime'))
    );

    CREATE TABLE IF NOT EXISTS resultados_taf (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      militar_id INTEGER NOT NULL,
      avaliacao_id INTEGER NOT NULL,
      chamada INTEGER NOT NULL DEFAULT 1,
      corrida_12min INTEGER,
      flexao_bracos INTEGER,
      abdominal_supra INTEGER,
      flexao_barra INTEGER,
      ppm_tempo TEXT,
      conceito_corrida TEXT,
      conceito_flexao_bracos TEXT,
      conceito_abdominal TEXT,
      conceito_barra TEXT,
      conceito_ppm TEXT,
      conceito_global TEXT,
      suficiencia TEXT,
      padrao_verificado TEXT,
      taf_alternativo INTEGER DEFAULT 0,
      observacoes TEXT,
      criado_em TEXT DEFAULT (datetime('now','localtime')),
      FOREIGN KEY (militar_id) REFERENCES militares(id),
      FOREIGN KEY (avaliacao_id) REFERENCES avaliacoes(id),
      UNIQUE(militar_id, avaliacao_id, chamada)
    );
  `);
}

// ── Criar janela ──────────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 1024,
    minHeight: 640,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
      webSecurity: false,
      allowRunningInsecureContent: false,
    },
    icon: path.join(__dirname, 'src/assets/img/icon.png'),
    title: 'TAF-EB — Sistema de Avaliação Física',
    show: false
  });

  mainWindow.loadFile('src/index.html');
  mainWindow.once('ready-to-show', () => mainWindow.show());
}

// ── Definir pasta de dados do usuário dentro do projeto ───────────────────
// Evita erros de permissão ao tentar escrever em pastas do sistema Windows
// ex: "Unable to move the cache: Acesso negado. (0x5)"
const userDataLocal = path.join(__dirname, 'dados');
if (!fs.existsSync(userDataLocal)) {
  try { fs.mkdirSync(userDataLocal, { recursive: true }); } catch(e) {}
}
app.setPath('userData', userDataLocal);
app.setPath('logs',     path.join(userDataLocal, 'logs'));

app.whenReady().then(async () => {
  await initDatabase();
  createWindow();

  // Permitir window.open para impressão de PDF/relatórios
  const { shell } = require('electron');
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    // URLs de blob (PDF interno) — abrir como nova janela Electron
    if (url.startsWith('about:blank') || url === '') {
      return {
        action: 'allow',
        overrideBrowserWindowOptions: {
          width: 1100, height: 800,
          webPreferences: { nodeIntegration: false, contextIsolation: false },
        }
      };
    }
    // URLs externas — abrir no navegador padrão
    shell.openExternal(url);
    return { action: 'deny' };
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ═══════════════════════════════════════════════════════════════════════════
// IPC HANDLERS
// ═══════════════════════════════════════════════════════════════════════════

// ── MILITARES ─────────────────────────────────────────────────────────────
ipcMain.handle('militares:listar', () => {
  const stmt = db.prepare(`SELECT * FROM militares ORDER BY COALESCE(nome_guerra, nome)`);
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
});

ipcMain.handle('militares:salvar', (_, dados) => {
  if (dados.id) {
    db.run(`UPDATE militares SET nome=?, nome_guerra=?, posto_graduacao=?, arma=?, lem=?,
            situacao_funcional=?, sexo=?, data_nascimento=?, om=?, subunidade=?
            WHERE id=?`,
      [dados.nome, dados.nome_guerra, dados.posto_graduacao, dados.arma, dados.lem,
       dados.situacao_funcional, dados.sexo, dados.data_nascimento,
       dados.om, dados.subunidade, dados.id]);
  } else {
    db.run(`INSERT INTO militares (nome, nome_guerra, posto_graduacao, arma, lem,
            situacao_funcional, sexo, data_nascimento, om, subunidade)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [dados.nome, dados.nome_guerra, dados.posto_graduacao, dados.arma, dados.lem,
       dados.situacao_funcional, dados.sexo, dados.data_nascimento,
       dados.om, dados.subunidade]);
  }
  salvarDb();
  return { ok: true };
});

ipcMain.handle('militares:excluir', (_, id) => {
  db.run(`UPDATE militares SET ativo = 0 WHERE id = ?`, [id]);
  salvarDb();
  return { ok: true };
});

ipcMain.handle('militares:buscar', (_, id) => {
  const stmt = db.prepare(`SELECT * FROM militares WHERE id = ?`);
  stmt.bind([id]);
  const row = stmt.step() ? stmt.getAsObject() : null;
  stmt.free();
  return row;
});

// ── AVALIAÇÕES ────────────────────────────────────────────────────────────
ipcMain.handle('avaliacoes:listar', () => {
  const stmt = db.prepare(`SELECT * FROM avaliacoes WHERE ativa = 1 ORDER BY ano DESC, id DESC`);
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
});

ipcMain.handle('avaliacoes:salvar', (_, dados) => {
  if (dados.id) {
    db.run(`UPDATE avaliacoes SET descricao=?, tipo=?, ano=?, padrao=?, situacao_funcional=?,
            data_1_chamada=?, data_2_chamada=?, data_3_chamada=?, data_chamada_extra=? WHERE id=?`,
      [dados.descricao, dados.tipo, dados.ano,
       dados.padrao || 'PAD', dados.situacao_funcional || 'OM_Op',
       dados.data_1_chamada, dados.data_2_chamada,
       dados.data_3_chamada, dados.data_chamada_extra, dados.id]);
  } else {
    db.run(`INSERT INTO avaliacoes (descricao, tipo, ano, padrao, situacao_funcional,
            data_1_chamada, data_2_chamada, data_3_chamada, data_chamada_extra)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [dados.descricao, dados.tipo, dados.ano,
       dados.padrao || 'PAD', dados.situacao_funcional || 'OM_Op',
       dados.data_1_chamada, dados.data_2_chamada,
       dados.data_3_chamada, dados.data_chamada_extra]);
  }
  salvarDb();
  return { ok: true };
});

ipcMain.handle('avaliacoes:excluir', (_, id) => {
  db.run(`UPDATE avaliacoes SET ativa = 0 WHERE id = ?`, [id]);
  salvarDb();
  return { ok: true };
});

// ── RESULTADOS TAF ────────────────────────────────────────────────────────
ipcMain.handle('resultados:salvar', (_, dados) => {
  db.run(`INSERT OR REPLACE INTO resultados_taf
    (militar_id, avaliacao_id, chamada, corrida_12min, flexao_bracos,
     abdominal_supra, flexao_barra, ppm_tempo, conceito_corrida,
     conceito_flexao_bracos, conceito_abdominal, conceito_barra, conceito_ppm,
     conceito_global, suficiencia, padrao_verificado, taf_alternativo, observacoes)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [dados.militar_id, dados.avaliacao_id, dados.chamada,
     dados.corrida_12min, dados.flexao_bracos, dados.abdominal_supra,
     dados.flexao_barra, dados.ppm_tempo,
     dados.conceito_corrida, dados.conceito_flexao_bracos, dados.conceito_abdominal,
     dados.conceito_barra, dados.conceito_ppm,
     dados.conceito_global, dados.suficiencia, dados.padrao_verificado,
     dados.taf_alternativo ? 1 : 0, dados.observacoes]);
  salvarDb();
  return { ok: true };
});

ipcMain.handle('resultados:listar', (_, avaliacao_id) => {
  const stmt = db.prepare(`
    SELECT r.*, m.nome, m.nome_guerra, m.posto_graduacao, m.sexo, m.arma,
           m.data_nascimento, m.lem, m.situacao_funcional, m.om, m.subunidade
    FROM resultados_taf r
    JOIN militares m ON r.militar_id = m.id
    JOIN avaliacoes a ON r.avaliacao_id = a.id
    WHERE r.avaliacao_id = ? AND a.ativa = 1
    ORDER BY COALESCE(m.nome_guerra, m.nome)
  `);
  stmt.bind([avaliacao_id]);
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
});

// Lista TODOS os militares com ou sem resultado — para painel de lançamento
ipcMain.handle('resultados:listarComMilitares', (_, avaliacao_id, chamada) => {
  const stmt = db.prepare(`
    SELECT m.id as militar_id, m.nome, m.nome_guerra, m.posto_graduacao, m.sexo, m.arma,
           m.data_nascimento, m.lem, m.situacao_funcional, m.om, m.subunidade,
           r.id as resultado_id, r.chamada, r.corrida_12min, r.flexao_bracos,
           r.abdominal_supra, r.flexao_barra, r.ppm_tempo,
           r.conceito_corrida, r.conceito_flexao_bracos, r.conceito_abdominal,
           r.conceito_barra, r.conceito_ppm, r.conceito_global, r.suficiencia,
           r.padrao_verificado, r.observacoes
    FROM militares m
    LEFT JOIN resultados_taf r
      ON r.militar_id = m.id AND r.avaliacao_id = ? AND r.chamada = ?
    ORDER BY COALESCE(m.nome_guerra, m.nome)
  `);
  stmt.bind([avaliacao_id, chamada]);
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
});

ipcMain.handle('resultados:porMilitar', (_, militar_id) => {
  const stmt = db.prepare(`
    SELECT r.*, a.descricao, a.tipo, a.ano,
           a.data_1_chamada, a.data_2_chamada, a.data_3_chamada, a.data_chamada_extra
    FROM resultados_taf r
    JOIN avaliacoes a ON r.avaliacao_id = a.id
    WHERE r.militar_id = ? AND a.ativa = 1
    ORDER BY a.ano ASC,
             CASE a.tipo
               WHEN 'diagnostica' THEN 0
               WHEN '1taf' THEN 1
               WHEN '2taf' THEN 2
               WHEN '3taf' THEN 3
               ELSE 4 END ASC,
             r.chamada ASC
  `);
  stmt.bind([militar_id]);
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
});

ipcMain.handle('resultados:buscar', (_, militar_id, avaliacao_id, chamada) => {
  const stmt = db.prepare(`SELECT * FROM resultados_taf WHERE militar_id=? AND avaliacao_id=? AND chamada=?`);
  stmt.bind([militar_id, avaliacao_id, chamada]);
  const row = stmt.step() ? stmt.getAsObject() : null;
  stmt.free();
  return row;
});

// ── BACKUP ────────────────────────────────────────────────────────────────
ipcMain.handle('db:backup', async () => {
  const { filePath } = await dialog.showSaveDialog({
    title: 'Salvar backup do banco de dados',
    defaultPath: `taf-eb-backup-${new Date().toISOString().slice(0,10)}.db`,
    filters: [{ name: 'Banco de Dados', extensions: ['db'] }]
  });
  if (filePath) {
    fs.copyFileSync(getDbPath(), filePath);
    return { ok: true, path: filePath };
  }
  return { ok: false };
});

ipcMain.handle('db:caminho', () => getDbPath());
