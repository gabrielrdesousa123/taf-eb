<div align="center">

# TAF-EB
### Sistema de Avaliação Física — Exército Brasileiro

**Desktop app para gestão e análise do Teste de Aptidão Física conforme**  
**Portaria EME/C Ex Nº 850/2022**

![Electron](https://img.shields.io/badge/Electron-47848F?style=for-the-badge&logo=electron&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![SQLite](https://img.shields.io/badge/SQLite-003B57?style=for-the-badge&logo=sqlite&logoColor=white)

</div>

---

## Sobre o projeto

O **TAF-EB** é um sistema desktop desenvolvido para facilitar o gerenciamento das avaliações físicas do Exército Brasileiro. Permite o cadastro de militares, lançamento de resultados por chamada, cálculo automático de menções e suficiência, e geração de relatórios em PDF e Word.

> ⚠️ **Este repositório contém dados fictícios para demonstração.**  
> Nunca suba arquivos `.db` com dados reais de militares.

---

## Funcionalidades

- ✅ Cadastro de militares (LEMB e LEMS, masculino e feminino)
- ✅ Gerenciamento de avaliações (AD, 1º TAF, 2º TAF, 3º TAF)
- ✅ Lançamento de resultados por chamada com cálculo automático
- ✅ Importação em lote via CSV
- ✅ Dashboard com gráficos de desempenho por OII e menção global
- ✅ Fichas individuais com histórico e evolução gráfica
- ✅ Exportação de fichas em PDF e Word
- ✅ Backup do banco de dados
- ✅ Funciona 100% offline — sem internet, sem servidor

---

## Motor de cálculo

O cálculo segue rigorosamente os **Anexos A, B, C e D** da Portaria EME/C Ex Nº 850/2022:

| Tabela | Descrição |
|---|---|
| Anexo A | LEMB Masculino |
| Anexo B | LEMB Feminino |
| Anexo C | LEMS Masculino |
| Anexo D | LEMS Feminino |

**OIIs avaliados:** Corrida 12 minutos · Flexão de braço · Abdominal · Barra fixa (LEMB ≤49 anos)

**Padrões:** PBD (mín. Regular) · PAD (mín. Bom) · PED (mín. Muito Bom)

---

## Requisitos

- **Windows** 10 ou superior
- **Node.js** 18+ → [nodejs.org](https://nodejs.org)
- **Git** (opcional, para clonar)

---

## Instalação

```bash
# 1. Clonar o repositório
git clone https://github.com/SEU_USUARIO/taf-eb.git
cd taf-eb

# 2. Instalar dependências
npm install

# 3. Instalar Electron globalmente
npm install -g electron

# 4. (Opcional) Carregar dados de demonstração
node seed-demo.js

# 5. Executar
electron .
```

**Ou no Windows:** clique duas vezes em `INSTALAR.bat` e depois `EXECUTAR.bat`

---

## Estrutura do projeto

```
taf-eb/
├── main.js              # Processo principal Electron + banco SQLite
├── preload.js           # Bridge segura renderer ↔ main
├── seed-demo.js         # Script para popular dados fictícios de demo
├── src/
│   ├── index.html       # Shell da aplicação
│   ├── taf-calculos.js  # Motor de cálculo TAF (Portaria 850/2022)
│   └── assets/
│       ├── js/
│       │   ├── app.js        # Toda a lógica de interface (~210kb)
│       │   └── jszip.min.js  # Geração de arquivos DOCX
│       └── css/
│           └── main.css      # Estilo da aplicação
├── dados/               # Banco SQLite (gerado em runtime — não versionado)
│   └── .gitkeep
├── .gitignore
├── INSTALAR.bat
├── EXECUTAR.bat
└── package.json
```

---

## Tecnologias

| Tecnologia | Uso |
|---|---|
| [Electron](https://electronjs.org) | Framework desktop |
| [sql.js](https://sql.js.org) | SQLite via WebAssembly |
| [JSZip](https://stuk.github.io/jszip/) | Geração de arquivos Word (.docx) |
| HTML5 Canvas | Gráficos de evolução nas fichas |
| Vanilla JS | Interface sem frameworks |

---

## Contribuindo

Contribuições são bem-vindas! Se encontrar um bug ou tiver sugestão:

1. Abra uma [Issue](../../issues)
2. Faça um fork e crie um branch: `git checkout -b feature/minha-feature`
3. Commit: `git commit -m 'feat: minha feature'`
4. Push: `git push origin feature/minha-feature`
5. Abra um Pull Request

---

## Aviso legal

Este software é um projeto independente e **não é oficial do Exército Brasileiro**.  
Desenvolvido com base na Portaria EME/C Ex Nº 850/2022, de acesso público.  
Os dados dos militares são de responsabilidade exclusiva do operador do sistema.

---

## Licença

MIT License — veja [LICENSE](LICENSE) para detalhes.

---

<div align="center">
Desenvolvido por <strong>Gabriel Sousa</strong> — Professor de Educação Física · Mestre em Saúde Coletiva
</div>
