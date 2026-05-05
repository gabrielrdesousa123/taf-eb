/**
 * seed-demo.js
 * Popula o banco com dados fictícios para demonstração.
 * Execute: node seed-demo.js
 */

const path = require('path');
const fs   = require('fs');

// Carregar sql.js
const initSqlJs = require('sql.js');

const DB_PATH = path.join(__dirname, 'dados', 'taf-eb.db');

async function seed() {
  const SQL = await initSqlJs();

  // Criar banco do zero
  if (fs.existsSync(DB_PATH)) fs.unlinkSync(DB_PATH);
  const db = new SQL.Database();

  // Criar tabelas
  db.run(`
    CREATE TABLE IF NOT EXISTS militares (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT NOT NULL, nome_guerra TEXT, posto_graduacao TEXT,
      arma TEXT, lem TEXT, situacao_funcional TEXT, sexo TEXT,
      data_nascimento TEXT, om TEXT, subunidade TEXT, ativo INTEGER DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS avaliacoes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      descricao TEXT NOT NULL, tipo TEXT, ano INTEGER, padrao TEXT,
      situacao_funcional TEXT,
      data_1_chamada TEXT, data_2_chamada TEXT,
      data_3_chamada TEXT, data_chamada_extra TEXT, ativa INTEGER DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS resultados_taf (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      militar_id INTEGER, avaliacao_id INTEGER, chamada INTEGER DEFAULT 1,
      corrida_12min INTEGER, flexao_bracos INTEGER,
      abdominal_supra INTEGER, flexao_barra INTEGER, ppm_tempo TEXT,
      conceito_corrida TEXT, conceito_flexao_bracos TEXT,
      conceito_abdominal TEXT, conceito_barra TEXT, conceito_ppm TEXT,
      conceito_global TEXT, suficiencia TEXT, padrao_verificado TEXT,
      taf_alternativo INTEGER DEFAULT 0, observacoes TEXT,
      UNIQUE(militar_id, avaliacao_id, chamada)
    );
  `);

  // Militares fictícios
  const militares = [
    ['CARLOS EDUARDO MENDES','MENDES','Cap','CAV','LEMB','Estb_Ens','M','1990-03-15','1 RCB','1 Esqd'],
    ['RAFAEL SOUZA LIMA','LIMA','1Ten','INF','LEMB','Estb_Ens','M','1995-07-22','1 RCB','2 Cia'],
    ['THIAGO ALVES COSTA','COSTA','2Ten','ART','LEMB','Estb_Ens','M','1998-11-05','1 GAC','Btria A'],
    ['PEDRO HENRIQUE ROCHA','ROCHA','ST','CAV','LEMB','Estb_Ens','M','1988-04-30','1 RCB','1 Esqd'],
    ['LUCAS FERREIRA SANTOS','SANTOS','Sgt','INF','LEMB','Estb_Ens','M','1993-09-12','1 RCB','2 Cia'],
    ['MARCOS VINICIUS CARVALHO','CARVALHO','Cb','CAV','LEMB','Estb_Ens','M','2001-02-28','1 RCB','1 Esqd'],
    ['JOAO PAULO SILVA','SILVA','Sd','INF','LEMB','Estb_Ens','M','2003-06-17','1 RCB','2 Cia'],
    ['ANA CAROLINA PEREIRA','PEREIRA','Cap','QCO','LEMS','Estb_Ens','F','1991-08-10','1 RCB','S4'],
    ['FERNANDA OLIVEIRA COSTA','OLIVEIRA','1Ten','QCO','LEMS','Estb_Ens','F','1996-12-03','1 RCB','S1'],
    ['JULIANA MARTINS SOUZA','MARTINS','ST','QCO','LEMS','Estb_Ens','F','1987-05-19','1 RCB','S4'],
    ['BRUNO NASCIMENTO LIMA','NASCIMENTO','Cap','ENG','LEMB','Estb_Ens','M','1989-01-25','1 BEC','Cia A'],
    ['GUILHERME ARAUJO MENDES','ARAUJO','1Ten','ENG','LEMB','Estb_Ens','M','1994-10-08','1 BEC','Cia B'],
    ['LEANDRO RIBEIRO SANTOS','RIBEIRO','Sgt','CAV','LEMB','Estb_Ens','M','1992-03-22','1 RCB','1 Esqd'],
    ['RODRIGO DIAS FERREIRA','DIAS','Cap','COM','LEMB','Estb_Ens','M','1990-07-14','1 RCB','S6'],
    ['ALEXANDRE MOURA COSTA','MOURA','2Ten','CAV','LEMB','Estb_Ens','M','1999-09-30','1 RCB','1 Esqd'],
  ];

  militares.forEach(m => {
    db.run(`INSERT INTO militares (nome,nome_guerra,posto_graduacao,arma,lem,situacao_funcional,sexo,data_nascimento,om,subunidade)
            VALUES (?,?,?,?,?,?,?,?,?,?)`, m);
  });

  // Avaliações
  db.run(`INSERT INTO avaliacoes (descricao,tipo,ano,padrao,data_1_chamada,ativa)
          VALUES ('Avaliacao Diagnostica 2026','diagnostica',2026,'PBD','2026-02-25',1)`);
  db.run(`INSERT INTO avaliacoes (descricao,tipo,ano,padrao,data_1_chamada,ativa)
          VALUES ('1o TAF 2026','1taf',2026,'PBD','2026-04-26',1)`);

  // Resultados fictícios — AD 2026
  const resultadosAD = [
    [1,1,1, 2750,33,78,12, 'B','E','E','E','B','S'],
    [2,1,1, 2450,28,65,0,  'I','B','B',null,'I','NS'],
    [3,1,1, 2850,35,72,14, 'B','MB','MB','E','B','S'],
    [4,1,1, 2600,30,68,10, 'R','B','B','MB','R','NS'],
    [5,1,1, 3050,40,80,16, 'MB','E','E','E','MB','S'],
    [6,1,1, 2200,22,55,6,  'I','R','I','B','I','NS'],
    [7,1,1, 2100,20,50,4,  'I','I','I','R','I','NS'],
    [8,1,1, 2700,32,0,0,   'B','MB',null,null,'B','S'],
    [9,1,1, 2500,28,0,0,   'R','B',null,null,'R','NS'],
    [10,1,1,2900,35,0,0,   'MB','E',null,null,'MB','S'],
    [11,1,1,2650,31,70,11, 'B','B','MB','E','B','S'],
    [12,1,1,2400,25,60,8,  'R','R','B','MB','R','NS'],
    [13,1,1,2800,36,75,13, 'B','E','E','E','B','S'],
    [14,1,1,2550,29,66,9,  'R','B','B','B','R','NS'],
    [15,1,1,3100,42,82,18, 'E','E','E','E','E','S'],
  ];

  resultadosAD.forEach(r => {
    db.run(`INSERT INTO resultados_taf
            (militar_id,avaliacao_id,chamada,corrida_12min,flexao_bracos,abdominal_supra,flexao_barra,
             conceito_corrida,conceito_flexao_bracos,conceito_abdominal,conceito_barra,
             conceito_global,suficiencia,padrao_verificado)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,'PBD')`, r);
  });

  // Resultados fictícios — 1o TAF 2026 (evolução em relação ao AD)
  const resultados1TAF = [
    [1,2,1, 2850,36,80,13, 'B','E','E','E','B','S'],
    [2,2,1, 2550,30,68,5,  'R','B','B','B','R','NS'],
    [3,2,1, 2950,38,75,15, 'MB','E','E','E','MB','S'],
    [4,2,1, 2700,33,72,12, 'B','B','MB','E','B','S'],
    [5,2,1, 3150,42,83,17, 'E','E','E','E','E','S'],
    [6,2,1, 2350,25,60,8,  'R','B','B','MB','R','NS'],
    [7,2,1, 2250,22,55,6,  'I','R','R','B','I','NS'],
    [8,2,1, 2800,35,0,0,   'B','E',null,null,'B','S'],
    [9,2,1, 2600,30,0,0,   'B','MB',null,null,'B','S'],
    [10,2,1,3000,38,0,0,   'MB','E',null,null,'MB','S'],
    [11,2,1,2750,34,73,12, 'B','MB','E','E','B','S'],
    [12,2,1,2500,28,63,9,  'R','B','B','MB','R','NS'],
    [13,2,1,2900,38,78,14, 'B','E','E','E','B','S'],
    [14,2,1,2650,32,70,11, 'B','B','MB','MB','B','S'],
    [15,2,1,3200,45,85,20, 'E','E','E','E','E','S'],
  ];

  resultados1TAF.forEach(r => {
    db.run(`INSERT INTO resultados_taf
            (militar_id,avaliacao_id,chamada,corrida_12min,flexao_bracos,abdominal_supra,flexao_barra,
             conceito_corrida,conceito_flexao_bracos,conceito_abdominal,conceito_barra,
             conceito_global,suficiencia,padrao_verificado)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,'PBD')`, r);
  });

  // Salvar banco
  const data = db.export();
  fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });
  fs.writeFileSync(DB_PATH, Buffer.from(data));
  db.close();

  console.log('✓ Banco de demonstração criado em: ' + DB_PATH);
  console.log('✓ ' + militares.length + ' militares fictícios inseridos');
  console.log('✓ 2 avaliações (AD 2026 + 1o TAF 2026)');
  console.log('✓ ' + (resultadosAD.length + resultados1TAF.length) + ' resultados inseridos');
  console.log('\nAgora execute: EXECUTAR.bat');
}

seed().catch(console.error);
