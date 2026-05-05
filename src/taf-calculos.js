/**
 * TAF-EB — Motor de Cálculo
 * Portaria EME/C Ex Nº 850, de 31 de agosto de 2022 (EB20-D-03.053)
 *
 * Conceitos: I=Insuficiente, R=Regular, B=Bom, MB=Muito Bom, E=Excelente
 * Padrões: PBD=mínimo Regular, PAD=mínimo Bom, PED=mínimo Muito Bom
 */

// ── Tabelas de referência (Anexos A-D da Portaria) ─────────────────────────

const TABELAS = {

  // ── ANEXO A — Masculino LEMB ─────────────────────────────────────────────
  LEMB_M: {
    corrida: [
      { min: 18, max: 21, I: [0,2599], R: [2600,2799], B: [2800,3149], MB: [3150,3199], E: [3200,99999] },
      { min: 22, max: 25, I: [0,2699], R: [2700,2849], B: [2850,3099], MB: [3100,3249], E: [3250,99999] },
      { min: 26, max: 29, I: [0,2599], R: [2600,2749], B: [2750,2999], MB: [3000,3149], E: [3150,99999] },
      { min: 30, max: 33, I: [0,2549], R: [2550,2649], B: [2650,2899], MB: [2900,3099], E: [3100,99999] },
      { min: 34, max: 37, I: [0,2449], R: [2450,2549], B: [2550,2799], MB: [2800,2949], E: [2950,99999] },
      { min: 38, max: 41, I: [0,2349], R: [2350,2449], B: [2450,2699], MB: [2700,2849], E: [2850,99999] },
      { min: 42, max: 45, I: [0,2249], R: [2250,2399], B: [2400,2599], MB: [2600,2749], E: [2750,99999] },
      { min: 46, max: 49, I: [0,2149], R: [2150,2299], B: [2300,2499], MB: [2500,2649], E: [2650,99999] },
      { min: 50, max: 53, suf: 1900 },
      { min: 54, max: 57, suf: 1800 },
      { min: 58, max: 61, suf: 1600 },
      { min: 62, max: 65, suf: 1400 },
    ],
    flexao_bracos: [
      { min: 18, max: 21, I: [0,21], R: [22,24], B: [25,33], MB: [34,38], E: [39,9999] },
      { min: 22, max: 25, I: [0,23], R: [24,26], B: [27,35], MB: [36,40], E: [41,9999] },
      { min: 26, max: 29, I: [0,21], R: [22,24], B: [25,33], MB: [34,38], E: [39,9999] },
      { min: 30, max: 33, I: [0,20], R: [21,23], B: [24,31], MB: [32,36], E: [37,9999] },
      { min: 34, max: 37, I: [0,17], R: [18,20], B: [21,28], MB: [29,33], E: [34,9999] },
      { min: 38, max: 41, I: [0,16], R: [17,19], B: [20,27], MB: [28,31], E: [32,9999] },
      { min: 42, max: 45, I: [0,14], R: [15,17], B: [18,24], MB: [25,28], E: [29,9999] },
      { min: 46, max: 49, I: [0,11], R: [12,14], B: [15,21], MB: [22,25], E: [26,9999] },
      { min: 50, max: 53, suf: 11 },
      { min: 54, max: 57, suf: 9 },
      { min: 58, max: 61, suf: 8 },
      { min: 62, max: 65, suf: 6 },
    ],
    abdominal: [
      { min: 18, max: 21, I: [0,34], R: [35,44], B: [45,63], MB: [64,73], E: [74,9999] },
      { min: 22, max: 25, I: [0,41], R: [42,51], B: [52,68], MB: [69,78], E: [79,9999] },
      { min: 26, max: 29, I: [0,37], R: [38,48], B: [49,65], MB: [66,75], E: [76,9999] },
      { min: 30, max: 33, I: [0,33], R: [34,42], B: [43,60], MB: [61,69], E: [70,9999] },
      { min: 34, max: 37, I: [0,30], R: [31,39], B: [40,56], MB: [57,65], E: [66,9999] },
      { min: 38, max: 41, I: [0,28], R: [29,37], B: [38,54], MB: [55,63], E: [64,9999] },
      { min: 42, max: 45, I: [0,26], R: [27,35], B: [36,52], MB: [53,61], E: [62,9999] },
      { min: 46, max: 49, I: [0,24], R: [25,33], B: [34,50], MB: [51,59], E: [60,9999] },
      { min: 50, max: 53, suf: 23 },
      { min: 54, max: 57, suf: 21 },
      { min: 58, max: 61, suf: 19 },
      { min: 62, max: 65, suf: 17 },
    ],
    barra: [
      { min: 18, max: 21, I: [0,4], R: [5,6], B: [7,9], MB: [10,11], E: [12,9999] },
      { min: 22, max: 25, I: [0,5], R: [6,7], B: [8,10], MB: [11,12], E: [13,9999] },
      { min: 26, max: 29, I: [0,4], R: [5,6], B: [7,9], MB: [10,11], E: [12,9999] },
      { min: 30, max: 33, I: [0,4], R: [5,5], B: [6,8], MB: [9,10], E: [11,9999] },
      { min: 34, max: 37, I: [0,3], R: [4,4], B: [5,6], MB: [7,8], E: [9,9999] },
      { min: 38, max: 39, I: [0,2], R: [3,3], B: [4,5], MB: [6,7], E: [8,9999] },
      { min: 40, max: 45, suf: 2 },
      { min: 46, max: 49, suf: 1 },
    ],
    // PPM — tempo em segundos (PED e PAD)
    ppm: [
      { min: 18, max: 21, PED: 310, PAD: 340 },
      { min: 22, max: 25, PED: 300, PAD: 330 },
      { min: 26, max: 29, PED: 330, PAD: 360 },
      { min: 30, max: 33, PED: 380, PAD: 410 },
      { min: 34, max: 37, PED: 410, PAD: 440 },
      { min: 38, max: 41, PED: 440, PAD: 470 },
    ],
  },

  // ── ANEXO B — Feminino LEMB ──────────────────────────────────────────────
  LEMB_F: {
    corrida: [
      { min: 18, max: 21, I: [0,2099], R: [2100,2199], B: [2200,2449], MB: [2450,2599], E: [2600,99999] },
      { min: 22, max: 25, I: [0,2149], R: [2150,2249], B: [2250,2449], MB: [2450,2649], E: [2650,99999] },
      { min: 26, max: 29, I: [0,2099], R: [2100,2199], B: [2200,2399], MB: [2400,2599], E: [2600,99999] },
      { min: 30, max: 33, I: [0,2049], R: [2050,2149], B: [2150,2349], MB: [2350,2549], E: [2550,99999] },
      { min: 34, max: 37, I: [0,1999], R: [2000,2099], B: [2100,2299], MB: [2300,2499], E: [2500,99999] },
      { min: 38, max: 41, I: [0,1899], R: [1900,1999], B: [2000,2199], MB: [2200,2399], E: [2400,99999] },
      { min: 42, max: 45, I: [0,1849], R: [1850,1949], B: [1950,2149], MB: [2150,2299], E: [2300,99999] },
      { min: 46, max: 49, I: [0,1749], R: [1750,1849], B: [1850,2049], MB: [2050,2249], E: [2250,99999] },
      { min: 50, max: 53, suf: 1600 },
      { min: 54, max: 57, suf: 1500 },
      { min: 58, max: 61, suf: 1300 },
      { min: 62, max: 65, suf: 1100 },
    ],
    flexao_bracos: [
      { min: 18, max: 21, I: [0,10], R: [11,11], B: [12,16], MB: [17,19], E: [20,9999] },
      { min: 22, max: 25, I: [0,11], R: [12,12], B: [13,18], MB: [19,21], E: [22,9999] },
      { min: 26, max: 29, I: [0,10], R: [11,11], B: [12,16], MB: [17,19], E: [20,9999] },
      { min: 30, max: 33, I: [0,9],  R: [10,10], B: [11,15], MB: [16,18], E: [19,9999] },
      { min: 34, max: 37, I: [0,8],  R: [9,9],   B: [10,14], MB: [15,17], E: [18,9999] },
      { min: 38, max: 41, I: [0,7],  R: [8,8],   B: [9,13],  MB: [14,16], E: [17,9999] },
      { min: 42, max: 45, I: [0,6],  R: [7,7],   B: [8,12],  MB: [13,15], E: [16,9999] },
      { min: 46, max: 49, I: [0,5],  R: [6,6],   B: [7,10],  MB: [11,13], E: [14,9999] },
      { min: 50, max: 53, suf: 5 },
      { min: 54, max: 57, suf: 4 },
      { min: 58, max: 61, suf: 3 },
      { min: 62, max: 65, suf: 2 },
    ],
    abdominal: [
      { min: 18, max: 21, I: [0,30], R: [31,39], B: [40,55], MB: [56,64], E: [65,9999] },
      { min: 22, max: 25, I: [0,32], R: [33,41], B: [42,57], MB: [58,66], E: [67,9999] },
      { min: 26, max: 29, I: [0,31], R: [32,40], B: [41,56], MB: [57,65], E: [66,9999] },
      { min: 30, max: 33, I: [0,29], R: [30,38], B: [39,54], MB: [55,63], E: [64,9999] },
      { min: 34, max: 37, I: [0,27], R: [28,36], B: [37,52], MB: [53,61], E: [62,9999] },
      { min: 38, max: 41, I: [0,25], R: [26,34], B: [35,50], MB: [51,59], E: [60,9999] },
      { min: 42, max: 45, I: [0,23], R: [24,32], B: [33,48], MB: [49,57], E: [58,9999] },
      { min: 46, max: 49, I: [0,21], R: [22,30], B: [31,46], MB: [47,55], E: [56,9999] },
      { min: 50, max: 53, suf: 20 },
      { min: 54, max: 57, suf: 18 },
      { min: 58, max: 61, suf: 16 },
      { min: 62, max: 65, suf: 14 },
    ],
    barra: [
      { min: 18, max: 21, I: [0,0], R: [1,2], B: [3,4], MB: [5,5], E: [6,9999] },
      { min: 22, max: 25, I: [0,1], R: [2,3], B: [4,5], MB: [6,6], E: [7,9999] },
      { min: 26, max: 29, I: [0,1], R: [2,2], B: [3,4], MB: [5,5], E: [6,9999] },
      { min: 30, max: 33, I: [0,0], R: [1,2], B: [3,4], MB: [5,5], E: [6,9999] },
      { min: 34, max: 37, I: [0,0], R: [1,2], B: [3,3], MB: [4,4], E: [5,9999] },
      { min: 38, max: 39, I: [0,0], R: [1,1], B: [2,2], MB: [3,3], E: [4,9999] },
      { min: 40, max: 45, suf_seg: 45 },   // sustentação em segundos
      { min: 46, max: 49, suf_seg: 30 },
    ],
    ppm: [ // mesmo que masculino (tabela idêntica no documento)
      { min: 18, max: 21, PED: 310, PAD: 340 },
      { min: 22, max: 25, PED: 300, PAD: 330 },
      { min: 26, max: 29, PED: 330, PAD: 360 },
      { min: 30, max: 33, PED: 380, PAD: 410 },
      { min: 34, max: 37, PED: 410, PAD: 440 },
      { min: 38, max: 41, PED: 440, PAD: 470 },
    ],
  },

  // ── ANEXO C — Masculino LEMS/LEMC/LEMCT ─────────────────────────────────
  LEMS_M: {
    corrida: [
      { min: 18, max: 21, I: [0,2599], R: [2600,2699], B: [2700,2899], MB: [2900,2999], E: [3000,99999] },
      { min: 22, max: 25, I: [0,2699], R: [2700,2799], B: [2800,2949], MB: [2950,3049], E: [3050,99999] },
      { min: 26, max: 29, I: [0,2599], R: [2600,2699], B: [2700,2849], MB: [2850,2949], E: [2950,99999] },
      { min: 30, max: 33, I: [0,2549], R: [2550,2649], B: [2650,2799], MB: [2800,2899], E: [2900,99999] },
      { min: 34, max: 37, I: [0,2449], R: [2450,2549], B: [2550,2649], MB: [2650,2749], E: [2750,99999] },
      { min: 38, max: 41, I: [0,2349], R: [2350,2449], B: [2450,2549], MB: [2550,2649], E: [2650,99999] },
      { min: 42, max: 45, I: [0,2249], R: [2250,2349], B: [2350,2499], MB: [2500,2599], E: [2600,99999] },
      { min: 46, max: 49, I: [0,2149], R: [2150,2299], B: [2300,2399], MB: [2400,2499], E: [2500,99999] },
      { min: 50, max: 53, suf: 1900 },
      { min: 54, max: 57, suf: 1800 },
      { min: 58, max: 61, suf: 1600 },
      { min: 62, max: 65, suf: 1400 },
    ],
    flexao_bracos: [
      { min: 18, max: 21, I: [0,21], R: [22,24], B: [25,29], MB: [30,32], E: [33,9999] },
      { min: 22, max: 25, I: [0,23], R: [24,26], B: [27,31], MB: [32,34], E: [35,9999] },
      { min: 26, max: 29, I: [0,21], R: [22,24], B: [25,29], MB: [30,32], E: [33,9999] },
      { min: 30, max: 33, I: [0,20], R: [21,23], B: [24,27], MB: [28,30], E: [31,9999] },
      { min: 34, max: 37, I: [0,17], R: [18,20], B: [21,25], MB: [26,28], E: [29,9999] },
      { min: 38, max: 41, I: [0,16], R: [17,19], B: [20,23], MB: [24,26], E: [27,9999] },
      { min: 42, max: 45, I: [0,14], R: [15,17], B: [18,21], MB: [22,24], E: [25,9999] },
      { min: 46, max: 49, I: [0,11], R: [12,14], B: [15,18], MB: [19,21], E: [22,9999] },
      { min: 50, max: 53, suf: 11 },
      { min: 54, max: 57, suf: 9 },
      { min: 58, max: 61, suf: 8 },
      { min: 62, max: 65, suf: 6 },
    ],
    abdominal: [
      // Idêntico ao LEMB_M conforme Portaria
      { min: 18, max: 21, I: [0,34], R: [35,44], B: [45,63], MB: [64,73], E: [74,9999] },
      { min: 22, max: 25, I: [0,41], R: [42,51], B: [52,68], MB: [69,78], E: [79,9999] },
      { min: 26, max: 29, I: [0,37], R: [38,48], B: [49,65], MB: [66,75], E: [76,9999] },
      { min: 30, max: 33, I: [0,33], R: [34,42], B: [43,60], MB: [61,69], E: [70,9999] },
      { min: 34, max: 37, I: [0,30], R: [31,39], B: [40,56], MB: [57,65], E: [66,9999] },
      { min: 38, max: 41, I: [0,28], R: [29,37], B: [38,54], MB: [55,63], E: [64,9999] },
      { min: 42, max: 45, I: [0,26], R: [27,35], B: [36,52], MB: [53,61], E: [62,9999] },
      { min: 46, max: 49, I: [0,24], R: [25,33], B: [34,50], MB: [51,59], E: [60,9999] },
      { min: 50, max: 53, suf: 23 },
      { min: 54, max: 57, suf: 21 },
      { min: 58, max: 61, suf: 19 },
      { min: 62, max: 65, suf: 17 },
    ],
  },

  // ── ANEXO D — Feminino LEMS/LEMC/LEMCT ───────────────────────────────────
  LEMS_F: {
    corrida: [
      { min: 18, max: 21, I: [0,2099], R: [2100,2199], B: [2200,2299], MB: [2300,2399], E: [2400,99999] },
      { min: 22, max: 25, I: [0,2149], R: [2150,2249], B: [2250,2349], MB: [2350,2449], E: [2450,99999] },
      { min: 26, max: 29, I: [0,2099], R: [2100,2199], B: [2200,2299], MB: [2300,2399], E: [2400,99999] },
      { min: 30, max: 33, I: [0,2049], R: [2050,2149], B: [2150,2249], MB: [2250,2349], E: [2350,99999] },
      { min: 34, max: 37, I: [0,1999], R: [2000,2099], B: [2100,2199], MB: [2200,2299], E: [2300,99999] },
      { min: 38, max: 41, I: [0,1899], R: [1900,1999], B: [2000,2149], MB: [2150,2249], E: [2250,99999] },
      { min: 42, max: 45, I: [0,1849], R: [1850,1949], B: [1950,2049], MB: [2050,2149], E: [2150,99999] },
      { min: 46, max: 49, I: [0,1749], R: [1750,1849], B: [1850,1949], MB: [1950,2049], E: [2050,99999] },
      { min: 50, max: 53, suf: 1600 },
      { min: 54, max: 57, suf: 1500 },
      { min: 58, max: 61, suf: 1300 },
      { min: 62, max: 65, suf: 1100 },
    ],
    flexao_bracos: [
      { min: 18, max: 21, I: [0,10], R: [11,11], B: [12,14], MB: [15,15], E: [16,9999] },
      { min: 22, max: 25, I: [0,11], R: [12,12], B: [13,15], MB: [16,16], E: [17,9999] },
      { min: 26, max: 29, I: [0,10], R: [11,11], B: [12,14], MB: [15,15], E: [16,9999] },
      { min: 30, max: 33, I: [0,9],  R: [10,10], B: [11,13], MB: [14,14], E: [15,9999] },
      { min: 34, max: 37, I: [0,8],  R: [9,9],   B: [10,12], MB: [13,13], E: [14,9999] },
      { min: 38, max: 41, I: [0,7],  R: [8,8],   B: [9,11],  MB: [12,12], E: [13,9999] },
      { min: 42, max: 45, I: [0,6],  R: [7,7],   B: [8,10],  MB: [11,11], E: [12,9999] },
      { min: 46, max: 49, I: [0,5],  R: [6,6],   B: [7,9],   MB: [10,10], E: [11,9999] },
      { min: 50, max: 53, suf: 5 },
      { min: 54, max: 57, suf: 4 },
      { min: 58, max: 61, suf: 3 },
      { min: 62, max: 65, suf: 2 },
    ],
    abdominal: [
      // Idêntico ao LEMB_F conforme Portaria
      { min: 18, max: 21, I: [0,30], R: [31,39], B: [40,55], MB: [56,64], E: [65,9999] },
      { min: 22, max: 25, I: [0,32], R: [33,41], B: [42,57], MB: [58,66], E: [67,9999] },
      { min: 26, max: 29, I: [0,31], R: [32,40], B: [41,56], MB: [57,65], E: [66,9999] },
      { min: 30, max: 33, I: [0,29], R: [30,38], B: [39,54], MB: [55,63], E: [64,9999] },
      { min: 34, max: 37, I: [0,27], R: [28,36], B: [37,52], MB: [53,61], E: [62,9999] },
      { min: 38, max: 41, I: [0,25], R: [26,34], B: [35,50], MB: [51,59], E: [60,9999] },
      { min: 42, max: 45, I: [0,23], R: [24,32], B: [33,48], MB: [49,57], E: [58,9999] },
      { min: 46, max: 49, I: [0,21], R: [22,30], B: [31,46], MB: [47,55], E: [56,9999] },
      { min: 50, max: 53, suf: 19 },
      { min: 54, max: 57, suf: 17 },
      { min: 58, max: 61, suf: 15 },
      { min: 62, max: 65, suf: 13 },
    ],
  },
};

// ── Utilitários ───────────────────────────────────────────────────────────

/** Calcula idade em anos completos a partir da data de nascimento */
function calcularIdade(dataNascimento, dataReferencia) {
  const nasc = new Date(dataNascimento + 'T12:00:00');
  // Usa dataReferencia (data da chamada do TAF) se fornecida, senão usa hoje
  // NUNCA deve usar hoje para cálculo de menção — apenas para exibição de idade atual
  const ref  = dataReferencia ? new Date(dataReferencia + 'T12:00:00') : new Date();
  let idade = ref.getFullYear() - nasc.getFullYear();
  const m = ref.getMonth() - nasc.getMonth();
  if (m < 0 || (m === 0 && ref.getDate() < nasc.getDate())) idade--;
  return idade;
}

/** Seleciona a chave da tabela com base em lem e sexo */
function chaveTabela(lem, sexo) {
  const isLEMB = lem === 'LEMB';
  const isMasc = sexo === 'M';
  if (isLEMB && isMasc)  return 'LEMB_M';
  if (isLEMB && !isMasc) return 'LEMB_F';
  if (!isLEMB && isMasc) return 'LEMS_M';
  return 'LEMS_F';
}

/** Encontra a linha de faixa etária */
function linhaFaixa(tabOII, idade) {
  return tabOII.find(f => idade >= f.min && idade <= f.max) || null;
}

/** Avalia um valor numérico contra a faixa etária — retorna conceito */
function avaliarOII(tabOII, idade, valor) {
  if (valor === null || valor === undefined || valor === '') return null;
  valor = Number(valor);
  const faixa = linhaFaixa(tabOII, idade);
  if (!faixa) return null;

  // Faixa de suficiência apenas (>=50 anos corrida/abd/flex, >=40 barra)
  if (faixa.suf !== undefined) {
    return { conceito: valor >= faixa.suf ? 'S' : 'NS', somenteApreciacao: true };
  }
  if (faixa.suf_seg !== undefined) {
    // sustentação em barra feminino
    return { conceito: valor >= faixa.suf_seg ? 'S' : 'NS', somenteApreciacao: true };
  }

  for (const conc of ['E', 'MB', 'B', 'R', 'I']) {
    const [minV, maxV] = faixa[conc];
    if (valor >= minV && valor <= maxV) return { conceito: conc, somenteApreciacao: false };
  }
  return { conceito: 'I', somenteApreciacao: false };
}

/** Avalia PPM (tempo em segundos, menor = melhor) */
function avaliarPPM(tabPPM, idade, tempoSeg, padrao) {
  if (tempoSeg === null || tempoSeg === undefined || tempoSeg === '') return null;
  tempoSeg = Number(tempoSeg);
  const faixa = linhaFaixa(tabPPM, idade);
  if (!faixa) return null;
  const limite = padrao === 'PED' ? faixa.PED : faixa.PAD;
  return { conceito: tempoSeg <= limite ? 'S' : 'NS', somenteApreciacao: true };
}

/** Mínimo de conceito exigido por padrão */
const MINIMO_PADRAO = { PBD: 'R', PAD: 'B', PED: 'MB' };
const ORDEM_CONCEITO = ['I', 'R', 'B', 'MB', 'E'];

function conceitoAtendePadrao(conceito, padrao) {
  if (!conceito || conceito === 'NS' || conceito === 'S') {
    // suficiência direta
    return conceito === 'S';
  }
  const minIdx = ORDEM_CONCEITO.indexOf(MINIMO_PADRAO[padrao]);
  const obtIdx = ORDEM_CONCEITO.indexOf(conceito);
  return obtIdx >= minIdx;
}

/** Conceito mais baixo da lista (excluindo nulos) */
function menorConceito(conceitos) {
  const validos = conceitos.filter(c => c && ORDEM_CONCEITO.includes(c));
  if (!validos.length) return null;
  return validos.reduce((menor, c) => {
    return ORDEM_CONCEITO.indexOf(c) < ORDEM_CONCEITO.indexOf(menor) ? c : menor;
  });
}

// ─────────────────────────────────────────────────────────────────────────
// FUNÇÃO PRINCIPAL
// ─────────────────────────────────────────────────────────────────────────
/**
 * Calcula todos os conceitos e suficiência do TAF de um militar.
 *
 * @param {object} p
 * @param {string}  p.dataNascimento  — 'YYYY-MM-DD'
 * @param {string}  p.lem             — 'LEMB' | 'LEMS' | 'LEMC' | 'LEMCT'
 * @param {string}  p.sexo            — 'M' | 'F'
 * @param {string}  p.situacaoFuncional — 'OM_NOp' | 'OM_Op' | 'OM_F_Emp_Estrt' | 'Estb_Ens'
 * @param {string}  p.postoGraduacao  — ex: 'Cap', 'Ten', '2Sgt' etc.
 * @param {number}  p.corrida         — metros
 * @param {number}  p.flexao          — repetições
 * @param {number}  p.abdominal       — repetições
 * @param {number}  p.barra           — repetições (ou segundos para fem ≥40)
 * @param {number}  p.ppmTempo        — segundos (null se não aplicável)
 * @param {boolean} p.tafAlternativo
 */
function calcularTAF(p) {
  // dataChamada: data real da chamada do TAF (YYYY-MM-DD)
  // Se não fornecida, usa anoTAF-01-01 como fallback (nunca usa hoje para cálculo de menção)
  // anoTAF pode ser passado como p.anoTAF ou extraído de p.dataChamada
  let dataRef = p.dataChamada || null;
  if (!dataRef && p.anoTAF) {
    // Usar 1º de janeiro do ano do TAF como fallback conservador
    dataRef = `${p.anoTAF}-01-01`;
  }
  // Se ainda null, usa hoje (último recurso — só para exibição de ficha atual)
  const idade = calcularIdade(p.dataNascimento, dataRef);
  const isLEMB = p.lem === 'LEMB';
  const chave = chaveTabela(p.lem, p.sexo);
  const tab = TABELAS[chave];

  // Padrão exigido (Portaria §2.4):
  // Fonte prioritária: p.padrao (padrão configurado na avaliação TAF)
  // Fallback: deriva da situação funcional do militar
  let padrao;
  if (p.padrao && ['PBD','PAD','PED'].includes(p.padrao)) {
    padrao = p.padrao;
  } else {
    // Derivar da situação funcional do militar
    if (p.situacaoFuncional === 'OM_F_Emp_Estrt' && isLEMB) padrao = 'PED';
    else if (p.situacaoFuncional === 'OM_Op' && isLEMB)     padrao = 'PAD';
    else                                                      padrao = 'PBD';
  }

  const resultado = { idade, padrao, conceitos: {}, suficiencias: {}, suficienciaGlobal: '', conceitoGlobal: '' };

  // ── Corrida ──────────────────────────────────────────────────────────────
  const rc = avaliarOII(tab.corrida, idade, p.corrida);
  if (rc) {
    resultado.conceitos.corrida = rc.somenteApreciacao ? null : rc.conceito;
    resultado.suficiencias.corrida = rc.somenteApreciacao ? rc.conceito : (conceitoAtendePadrao(rc.conceito, padrao) ? 'S' : 'NS');
  }

  // ── Flexão de braços ─────────────────────────────────────────────────────
  const rf = avaliarOII(tab.flexao_bracos, idade, p.flexao);
  if (rf) {
    resultado.conceitos.flexao = rf.somenteApreciacao ? null : rf.conceito;
    resultado.suficiencias.flexao = rf.somenteApreciacao ? rf.conceito : (conceitoAtendePadrao(rf.conceito, padrao) ? 'S' : 'NS');
  }

  // ── Abdominal ────────────────────────────────────────────────────────────
  const ra = avaliarOII(tab.abdominal, idade, p.abdominal);
  if (ra) {
    resultado.conceitos.abdominal = ra.somenteApreciacao ? null : ra.conceito;
    resultado.suficiencias.abdominal = ra.somenteApreciacao ? ra.conceito : (conceitoAtendePadrao(ra.conceito, padrao) ? 'S' : 'NS');
  }

  // ── Barra (somente LEMB) ─────────────────────────────────────────────────
  if (isLEMB && tab.barra && idade <= 49) {
    const rb = avaliarOII(tab.barra, idade, p.barra);
    if (rb) {
      resultado.conceitos.barra = rb.somenteApreciacao ? null : rb.conceito;
      resultado.suficiencias.barra = rb.somenteApreciacao ? rb.conceito : (conceitoAtendePadrao(rb.conceito, padrao) ? 'S' : 'NS');
    }
  }

  // ── PPM (somente LEMB, PAD e PED, ≤ capitão/2Sgt, até 41 anos) ──────────
  if (isLEMB && tab.ppm && (padrao === 'PAD' || padrao === 'PED') && idade <= 41 && p.ppmTempo !== null) {
    const rp = avaliarPPM(tab.ppm, idade, p.ppmTempo, padrao);
    if (rp) {
      resultado.suficiencias.ppm = rp.conceito; // 'S' ou 'NS'
    }
  }

  // ── Conceito global ──────────────────────────────────────────────────────
  const conceitosValidos = Object.values(resultado.conceitos).filter(c => c && ORDEM_CONCEITO.includes(c));
  resultado.conceitoGlobal = menorConceito(conceitosValidos) || '—';

  // ── Suficiência global (contagiante — pior índice determina) ────────────
  const sufValores = Object.values(resultado.suficiencias);
  resultado.suficienciaGlobal = sufValores.includes('NS') ? 'NS' : (sufValores.length > 0 ? 'S' : '—');

  return resultado;
}

/** Converte "MM:SS" → segundos */
function mmssParaSegundos(str) {
  if (!str) return null;
  const partes = String(str).split(':');
  if (partes.length === 2) return parseInt(partes[0]) * 60 + parseInt(partes[1]);
  return parseInt(str);
}

// Expor para o browser (uso direto em renderer)
if (typeof window !== 'undefined') {
  window.TAF = { calcularTAF, calcularIdade, mmssParaSegundos, ORDEM_CONCEITO };
}
// Expor para Node/Electron main (se necessário)
if (typeof module !== 'undefined') {
  module.exports = { calcularTAF, calcularIdade, mmssParaSegundos, ORDEM_CONCEITO };
}
