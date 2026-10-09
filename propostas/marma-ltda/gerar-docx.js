// Gera as propostas em Word (.docx) a partir do mesmo conteúdo das versões HTML em português.
// Uso: node gerar-docx.js   (a partir da pasta propostas/marma-ltda)
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, LevelFormat, Footer, VerticalAlign, TableLayoutType,
} = require("docx");

const ASSETS = path.join(__dirname, "..", "..", "assets");
const INK = "1A1A1A", GREEN = "00AF58", GREEN_DARK = "008A45", GREEN_SOFT = "E8F7EF", LINE = "E3E5E8",
  SOFT = "F5F6F7", MUTED = "6B6B70", WARN = "B7791F", WARN_SOFT = "FDF6E9", GRAPHITE = "3C3A3D";
const W = 9638; // largura útil A4 com margens de 2 cm
const FONT = "Calibri";

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none };
const line = (c = LINE, s = 6) => ({ style: BorderStyle.SINGLE, size: s, color: c });
const allBorders = (c = LINE) => ({ top: line(c), bottom: line(c), left: line(c), right: line(c) });

// Texto com **negrito**
function runs(text, o = {}) {
  return text.split("**").map((t, i) => new TextRun({
    text: t, bold: o.bold || i % 2 === 1, color: o.color || INK, size: o.size || 20, font: FONT,
    italics: o.italics, characterSpacing: o.spacing,
  }));
}
const P = (text, o = {}) => new Paragraph({
  children: runs(text, o), alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 120, line: 276 },
});
const bullet = (text, o = {}) => new Paragraph({
  children: runs(text, { size: o.size || 19 }), numbering: { reference: "bul", level: 0 },
  spacing: { after: 60, line: 264 },
});
const small = (text) => P(text, { size: 16, color: MUTED });

function heading(n, text) {
  return new Paragraph({
    children: [
      new TextRun({ text: ` ${n} `, bold: true, color: "FFFFFF", size: 22, font: FONT, shading: { type: ShadingType.CLEAR, fill: GREEN, color: "auto" } }),
      new TextRun({ text: `  ${text}`, bold: true, color: INK, size: 26, font: FONT }),
    ],
    border: { bottom: line(LINE, 8) }, spacing: { before: 320, after: 160 }, keepNext: true,
  });
}

function cell(children, o = {}) {
  return new TableCell({
    children, width: { size: o.w, type: WidthType.DXA }, borders: o.borders || noBorders,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: "auto" } : undefined,
    margins: o.margins || { top: 140, bottom: 140, left: 200, right: 200 }, verticalAlign: o.valign,
    columnSpan: o.span,
  });
}
function table(cols, rows, o = {}) {
  return new Table({
    width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, rows, borders: o.borders,
  });
}
const spacer = (after = 120, keep = false) => new Paragraph({ children: [], spacing: { after }, keepNext: keep });

// Cartão: rótulo verde, título com etiquetas, lista
function tag(name) {
  const mb = name === "Metabase";
  return new TextRun({ text: `  ${name.toUpperCase()}  `, bold: true, size: 14, font: FONT, color: mb ? GREEN_DARK : GRAPHITE,
    shading: { type: ShadingType.CLEAR, fill: mb ? GREEN_SOFT : "ECECEE", color: "auto" } });
}
function cardContent(c) {
  const title = [new TextRun({ text: c.title + " ", bold: true, size: 21, font: FONT, color: INK })];
  (c.tags || []).forEach((t) => { title.push(new TextRun({ text: " ", size: 14 })); title.push(tag(t)); });
  return [
    new Paragraph({ children: [new TextRun({ text: c.step.toUpperCase(), bold: true, size: 15, color: GREEN_DARK, font: FONT, characterSpacing: 30 })], spacing: { after: 40 } }),
    new Paragraph({ children: title, spacing: { after: 100 } }),
    ...c.items.map((t) => bullet(t)),
  ];
}
function cards(list) {
  const gap = 200, half = (W - gap) / 2, rows = [];
  for (let i = 0; i < list.length; i += 2) {
    const a = list[i], b = list[i + 1];
    if (!b) {
      rows.push(new TableRow({ cantSplit: false, children: [cell(cardContent(a), { w: W, span: 3, borders: allBorders() })] }));
    } else {
      rows.push(new TableRow({ children: [
        cell(cardContent(a), { w: half, borders: allBorders() }),
        cell([new Paragraph("")], { w: gap }),
        cell(cardContent(b), { w: half, borders: allBorders() }),
      ] }));
    }
    rows.push(new TableRow({ children: [cell([new Paragraph({ children: [], spacing: { after: 0 } })], { w: W, span: 3, margins: { top: 40, bottom: 40, left: 0, right: 0 } })] }));
  }
  return table([half, gap, half], rows);
}

function box(text, warn) {
  return table([W], [new TableRow({ children: [cell([P(text, { size: 18, after: 0 })], {
    w: W, fill: warn ? WARN_SOFT : SOFT,
    borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 24, color: warn ? WARN : GREEN } },
  })] })]);
}

// Tabela de dados: cabeçalho preto, linha total verde-clara
function dataTable(cols, header, body, total, numCols = []) {
  // keepNext em todas as linhas menos a última mantém a tabela inteira na mesma página
  const mk =(vals, kind, last) => new TableRow({ cantSplit: true, tableHeader: kind === "h", children: vals.map((v, i) => cell([new Paragraph({
    children: runs(String(v), { size: kind === "h" ? 17 : 18, bold: kind !== "b", color: kind === "h" ? "FFFFFF" : INK }),
    alignment: numCols.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT, keepNext: !last,
  })], { w: cols[i], fill: kind === "h" ? INK : kind === "t" ? GREEN_SOFT : undefined,
    borders: { top: none, left: none, right: none, bottom: line() }, margins: { top: 90, bottom: 90, left: 160, right: 160 } })) });
  const rows = [mk(header, "h"), ...body.map((r, i) => mk(r, "b", !total && i === body.length - 1))];
  if (total) rows.push(mk(total, "t", true));
  return table(cols, rows);
}

function logos(numero) {
  const mbH = 70, mbW = Math.round(mbH * 823 / 475), acH = 84, acW = Math.round(acH * 1141 / 893);
  const img = (f, w, h) => new ImageRun({ type: "png", data: fs.readFileSync(path.join(ASSETS, f)), transformation: { width: w, height: h } });
  return table([2600, 2600, 4438], [new TableRow({ children: [
    cell([new Paragraph({ children: [img("logo-metabase.png", mbW, mbH)] })], { w: 2600, valign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 200 } }),
    cell([new Paragraph({ children: [img("logo-arnessen-cintra.png", acW, acH)] })], { w: 2600, valign: VerticalAlign.CENTER,
      borders: { top: none, bottom: none, right: none, left: line(LINE, 8) }, margins: { top: 0, bottom: 0, left: 300, right: 0 } }),
    cell([
      new Paragraph({ alignment: AlignmentType.RIGHT, children: runs(`**Proposta nº ${numero}**`, { size: 19 }), spacing: { after: 20 } }),
      new Paragraph({ alignment: AlignmentType.RIGHT, children: runs("Contabilidade + Advocacia em parceria", { size: 16, color: MUTED }) }),
    ], { w: 4438, valign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 0 } }),
  ] })]);
}

function hero(title, subtitle) {
  return table([W], [new TableRow({ children: [cell([
    new Paragraph({ children: [new TextRun({ text: "PROPOSTA COMERCIAL", bold: true, size: 15, color: GREEN, font: FONT, characterSpacing: 40 })], spacing: { after: 60 } }),
    new Paragraph({ children: [new TextRun({ text: title, bold: true, size: 40, color: "FFFFFF", font: FONT })], spacing: { after: 60 } }),
    new Paragraph({ children: [new TextRun({ text: subtitle, size: 20, color: "C9CBCF", font: FONT })] }),
  ], { w: W, fill: INK, margins: { top: 300, bottom: 300, left: 400, right: 300 },
    borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 48, color: GREEN } } })] })]);
}

function meta(items) {
  const w = W / 3;
  return table([w, w, w], [new TableRow({ children: items.map((it, i) => cell([
    new Paragraph({ children: [new TextRun({ text: it[0].toUpperCase(), size: 14, color: MUTED, font: FONT, characterSpacing: 30 })], spacing: { after: 20 } }),
    new Paragraph({ children: runs(`**${it[1]}**`, { size: 20 }), spacing: { after: 10 } }),
    new Paragraph({ children: runs(it[2], { size: 17, color: MUTED }) }),
  ], { w, borders: { top: none, left: none, bottom: line(), right: i < 2 ? line() : none }, margins: { top: 140, bottom: 140, left: i ? 240 : 0, right: 120 } })) })]);
}

function kpis(list) {
  const gap = 160, w = (W - 2 * gap) / 3, ch = [];
  list.forEach((k, i) => {
    const dark = i === 0;
    ch.push(cell([
      new Paragraph({ children: [new TextRun({ text: k[0], bold: true, size: 34, font: FONT, color: dark ? GREEN : INK })], spacing: { after: 40 } }),
      new Paragraph({ children: runs(k[1], { size: 16, color: dark ? "C9CBCF" : MUTED }) }),
    ], { w, fill: dark ? INK : SOFT, margins: { top: 180, bottom: 180, left: 220, right: 160 } }));
    if (i < 2) ch.push(cell([new Paragraph("")], { w: gap }));
  });
  return table([w, gap, w, gap, w], [new TableRow({ children: ch })]);
}

function bars(title, list, cap) {
  const lw = 2400, vw = 1300, bw = W - lw - vw - 400;
  const rows = [new TableRow({ children: [cell([P(`**${title}**`, { size: 20, after: 60 })], { w: W - 400, span: 3, margins: { top: 160, bottom: 0, left: 0, right: 0 } })] })];
  list.forEach(([label, sub, pct, val, green]) => {
    const fillW = Math.round(bw * pct);
    const barCols = pct < 1 ? [fillW, bw - fillW] : [bw];
    const barCells = [cell([new Paragraph({ children: [new TextRun({ text: " ", size: 14 })] })], { w: fillW, fill: green ? GREEN : "9A9A9F", margins: { top: 60, bottom: 60, left: 0, right: 0 } })];
    if (pct < 1) barCells.push(cell([new Paragraph({ children: [new TextRun({ text: " ", size: 14 })] })], { w: bw - fillW, fill: SOFT, margins: { top: 60, bottom: 60, left: 0, right: 0 } }));
    rows.push(new TableRow({ children: [
      cell([P(`**${label}**`, { size: 18, after: 0 }), P(sub, { size: 15, color: MUTED, after: 0 })], { w: lw, valign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 0, right: 100 } }),
      cell([table(barCols, [new TableRow({ children: barCells })])], { w: bw, valign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 0, right: 0 } }),
      cell([P(`**${val}**`, { size: 19, after: 0, align: AlignmentType.RIGHT })], { w: vw, valign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 100, right: 0 } }),
    ] }));
  });
  rows.push(new TableRow({ children: [cell([P(cap, { size: 16, color: MUTED, after: 0 })], { w: W - 400, span: 3, margins: { top: 80, bottom: 160, left: 0, right: 0 } })] }));
  return table([W], [new TableRow({ children: [cell([table([lw, bw, vw], rows)], { w: W, borders: allBorders(), margins: { top: 0, bottom: 0, left: 200, right: 200 } })] })]);
}

function timeline(list) {
  const w = W / list.length;
  return table(list.map(() => w), [new TableRow({ children: list.map(([a, b]) => cell([
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "●", size: 32, color: GREEN, font: FONT })], spacing: { after: 20 } }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: runs(`**${a}**`, { size: 20 }), spacing: { after: 10 } }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: runs(b, { size: 17, color: MUTED }) }),
  ], { w, borders: { top: line(GREEN_SOFT, 18), bottom: none, left: none, right: none } })) })]);
}

function priceCard(p) {
  const d = p.dark;
  const ch = [
    new Paragraph({ children: [new TextRun({ text: p.label.toUpperCase(), bold: true, size: 15, font: FONT, color: d ? GREEN : MUTED, characterSpacing: 30 })], spacing: { after: 60 } }),
    new Paragraph({ children: [new TextRun({ text: p.big, bold: true, size: p.bigSize || 46, font: FONT, color: d ? GREEN : INK })], spacing: { after: 80 } }),
    P(p.sub, { size: 17, color: d ? "C9CBCF" : MUTED, after: 100 }),
  ];
  if (p.split) ch.push(new Paragraph({ children: p.split.flatMap((s, i) => [
    ...(i ? [new TextRun({ text: "   ", size: 16 })] : []),
    new TextRun({ text: `  ${s}  `, size: 16, font: FONT, color: d ? "E6E6E6" : INK, shading: { type: ShadingType.CLEAR, fill: d ? "2C2C2E" : SOFT, color: "auto" } }),
  ]) }));
  return ch;
}
function prices(a, b) {
  const gap = 200, half = (W - gap) / 2;
  return table([half, gap, half], [new TableRow({ children: [
    cell(priceCard(a), { w: half, fill: a.dark ? INK : undefined, borders: a.dark ? undefined : allBorders(), margins: { top: 240, bottom: 240, left: 280, right: 240 } }),
    cell([new Paragraph("")], { w: gap }),
    cell(priceCard(b), { w: half, fill: b.dark ? INK : undefined, borders: b.dark ? undefined : allBorders(), margins: { top: 240, bottom: 240, left: 280, right: 240 } }),
  ] })]);
}

function signatures() {
  const gap = 300, w = (W - 2 * gap) / 3;
  const s = (a, b) => cell([P(`**${a}**`, { size: 18, align: AlignmentType.CENTER, after: 0 }), P(b, { size: 15, color: MUTED, align: AlignmentType.CENTER, after: 0 })],
    { w, borders: { top: line(INK, 8), bottom: none, left: none, right: none }, margins: { top: 80, bottom: 0, left: 0, right: 0 } });
  const g = () => cell([new Paragraph("")], { w: gap });
  return table([w, gap, w, gap, w], [new TableRow({ cantSplit: true, children: [
    s("MARMA LTDA", "Contratante"), g(),
    s("Estêvão Fernandes", "Contador · Metabase Contabilidade"), g(),
    s("Arnessen Cintra", "Advogado · OAB/PE 70.236"),
  ] })]);
}

// Bloco final indivisível: título, lista de responsabilidades e assinaturas ficam na mesma página
function closing(n, title, items) {
  return table([W], [new TableRow({ cantSplit: true, children: [cell([
    heading(n, title), ...items.map((t) => bullet(t, { size: 20 })), spacer(600), signatures(),
  ], { w: W, margins: { top: 0, bottom: 0, left: 0, right: 0 } })] })]);
}

function doc(children) {
  return new Document({
    creator: "Metabase Contabilidade",
    styles: { default: { document: { run: { font: FONT, size: 20, color: INK } } } },
    numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 300, hanging: 220 } }, run: { color: GREEN } } }] }] },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1134, right: 1134, footer: 500 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, border: { top: line(LINE, 6) }, spacing: { before: 100 }, children: [
        new TextRun({ text: "Metabase Contabilidade", bold: true, color: GREEN_DARK, size: 15, font: FONT }),
        new TextRun({ text: " · desde 2015 · 11 anos     |     ", color: MUTED, size: 15, font: FONT }),
        new TextRun({ text: "Arnessen Cintra Advocacia", bold: true, color: GRAPHITE, size: 15, font: FONT }),
        new TextRun({ text: " · OAB/PE 70.236", color: MUTED, size: 15, font: FONT }),
      ] })] }) },
      children,
    }],
  });
}

const cabecalho = (numero, titulo, sub, validade2) => [
  logos(numero), spacer(200), hero(titulo, sub),
  meta([["Cliente", "MARMA LTDA", "CNPJ 67.699.385/0001-06"], ["Data", "09/10/2026", "Santa Cruz do Capibaribe/PE"], ["Validade", "15 dias", validade2]]),
];
const juridicoNota = "Cada profissional emite seu próprio contrato e documento fiscal pela sua parte. A divisão segue as regras do CFC, para a contabilidade, e do Estatuto da Advocacia e Código de Ética da OAB, para a advocacia.";

// ---------- Proposta 001 ----------
const p1 = doc([
  ...cabecalho("001/2026", "Planejamento Trabalhista", "Diagnóstico, estimativa de custos e plano de regularização do quadro de empregados", "Prazo de execução: 30 dias"),
  heading(1, "Contexto"),
  P("A MARMA LTDA é uma indústria de confecção com CNPJ recém-aberto e ainda sem movimentação. Hoje cerca de **35 pessoas** trabalham nas instalações da empresa: parte recebe **por hora** e parte **por produção**. Quem trabalha por produção define o quanto quer produzir e, para cumprir as entregas e ganhar mais, costuma ter uma rotina intensa. A jornada de cada função será analisada caso a caso.", { size: 21 }),
  P("A empresa quer começar a operar formalmente, registrar os trabalhadores e reduzir o risco de **fiscalização do trabalho** e de **reclamações trabalhistas**, com um custo que caiba na operação. O foco principal são os trabalhadores **por produção**."),
  heading(2, "Objetivo"),
  P("Entregar à diretoria um **plano de regularização completo**, com o custo real de cada cenário e o modelo de contratação juridicamente seguro mais econômico para cada grupo de trabalhadores. O plano também traz o cronograma e os documentos prontos para começar os registros."),
  heading(3, "O que será entregue"),
  cards([
    { step: "Etapa 1", title: "Diagnóstico", tags: ["Metabase", "Arnessen"], items: [
      "Levantamento individual dos ~35 trabalhadores: função, forma de pagamento, remuneração média, jornada real, tempo de serviço e documentação.",
      "Visita técnica à fábrica e entrevistas sobre o fluxo de produção, o controle de peças e o horário de funcionamento.",
      "Enquadramento: CNAE, Simples Nacional (Anexo II), sindicato e Convenção Coletiva aplicável (piso, adicionais, regras de jornada e de produção)."] },
    { step: "Etapa 2", title: "Estudo de custos", tags: ["Metabase"], items: [
      "Custo mensal e anual de cada trabalhador e da folha: salário, FGTS, 13º, férias + 1/3, provisões, DSR, horas extras e saúde e segurança do trabalho.",
      "Comparação de cenários: registro integral × escalonado, e horista × produção com piso × banco de horas.",
      "Impacto no caixa e no custo de mão de obra por peça, e verificação dos limites do Simples e do ICMS-PE."] },
    { step: "Etapa 3", title: "Modelagem jurídica", tags: ["Arnessen"], items: [
      "**Contrato por produção** com garantia do piso, critério de apuração das peças, DSR e horas extras (art. 78 da CLT; art. 7º, VII, da CF; OJ 235 do TST).",
      "Jornada: análise caso a caso de cada função, acordo de compensação/banco de horas (art. 59 da CLT) e organização da escala semanal.",
      "Alternativas e seus riscos: prêmios (art. 457), contrato de experiência, facção externa (Lei 6.019/74) e os limites da \"pejotização\".",
      "**Passivo anterior ao registro** e estratégia para tratá-lo, inclusive por acordo extrajudicial homologado (art. 855-B da CLT).",
      "Regulamento interno e termos individuais."] },
    { step: "Etapa 4", title: "Plano de implantação", tags: ["Metabase", "Arnessen"], items: [
      "Cronograma de admissões por fases, com prioridades definidas pelo risco.",
      "Checklist admissional, eSocial, FGTS Digital e DCTFWeb.",
      "**Controle de ponto**, obrigatório acima de 20 empregados (art. 74, §2º, da CLT).",
      "Roteiro de saúde e segurança: PGR, PCMSO, exames, NR-12 (máquinas) e NR-17 (ergonomia)."] },
    { step: "Etapa 5", title: "Entrega e apresentação", tags: ["Metabase", "Arnessen"], items: [
      "Relatório executivo com os cenários, os custos e a recomendação final.",
      "Reunião presencial de apresentação à diretoria, com a participação do contador e do advogado."] },
  ]),
  box("**Fora do escopo:** execução mensal da folha e das admissões (Proposta nº 002/2026), defesa em ações judiciais e autos de infração já existentes, e custos de terceiros: exames médicos, PGR/PCMSO, sistema de ponto, taxas e certidões.", true),
  heading(4, "Estimativa preliminar de custos"),
  P("Premissas: salário-base de R$ 1.621,00 (salário mínimo 2026) e empresa optante pelo Simples Nacional, Anexo II. Nesse regime, a contribuição patronal ao INSS já está incluída no DAS. O piso da Convenção Coletiva ainda será confirmado e pode ser maior."),
  kpis([["R$ 2.153", "custo mensal por empregado registrado"], ["1,33×", "fator sobre o salário (encargos + provisões)"], ["R$ 75 mil", "folha mensal com os 35 registrados, sem horas extras"]]),
  spacer(160),
  dataTable([7138, 2500], ["Item mensal por empregado", "R$"], [
    ["Salário-base", "1.621,00"], ["FGTS (8%)", "129,68"], ["Provisão de 13º salário", "135,08"], ["Provisão de férias + 1/3", "180,11"],
    ["FGTS sobre 13º e férias", "25,22"], ["Provisão da multa do FGTS (40%)", "61,96"]], ["Custo mensal por empregado", "2.153,05"], [1]),
  spacer(80),
  dataTable([4638, 2500, 2500], ["Empregados registrados", "Folha/mês (R$)", "Encargos (R$)"], [
    ["10 registrados", "21.530", "5.320"], ["20 registrados", "43.061", "10.641"], ["35 registrados", "75.357", "18.622"]], null, [1, 2]),
  spacer(160),
  bars("Onde está a economia: custo de cada hora extra", [
    ["Horista", "hora + 50% + DSR", 1, "≈ R$ 17,60", false],
    ["Por produção", "só o adicional de 50% + DSR", 0.33, "≈ R$ 5,90", true]],
    "Valores por hora extra, com reflexos e encargos, sobre o salário-base. Quem ganha por produção recebe só o adicional sobre as horas extras, porque as peças já remuneram a hora (OJ 235 do TST). A economia chega a cerca de 2/3 do custo."),
  spacer(80),
  small("Números preliminares, só para dimensionar o trabalho. O estudo da Etapa 2 vai usar os dados reais de cada trabalhador, o piso da Convenção Coletiva e a jornada analisada caso a caso."),
  heading(5, "Prazo"),
  timeline([["Dia 0", "Assinatura e envio das informações"], ["Até o dia 10", "Diagnóstico concluído"], ["Até o dia 30", "Relatório final e apresentação"]]),
  heading(6, "Investimento"),
  prices(
    { label: "Planejamento completo", big: "R$ 4.800", sub: "50% na assinatura e 50% na entrega, ou 3× de R$ 1.600,00", split: ["Metabase: R$ 2.400", "Arnessen: R$ 2.400"] },
    { dark: true, label: "Condição especial", big: "R$ 4.000", sub: "Contratando também o acompanhamento mensal (Proposta nº 002/2026) até a entrega do relatório. As admissões iniciais ficam isentas da taxa de implantação.", split: ["Metabase: R$ 2.000", "Arnessen: R$ 2.000"] }),
  spacer(80),
  small(juridicoNota),
  closing(7, "Responsabilidades da contratante", [
    "Fornecer a relação dos trabalhadores e os dados de pagamento, produção e jornada dos últimos meses.",
    "Permitir o acesso à fábrica para a visita técnica e indicar um responsável interno.",
    "Prestar informações verdadeiras e completas. As recomendações dependem delas."
  ]),
]);

// ---------- Proposta 002 ----------
const p2 = doc([
  ...cabecalho("002/2026", "Acompanhamento Mensal", "Contabilidade, Departamento Pessoal e Assessoria Jurídica Trabalhista", "Vigência: prazo indeterminado"),
  heading(1, "Objetivo"),
  P("Depois do planejamento, a MARMA precisa de acompanhamento contínuo. A empresa vai crescer o quadro de empregados por etapas e deve manter a folha, os impostos e as rotinas trabalhistas em dia, com orientação jurídica preventiva para não criar passivo. A **Metabase Contabilidade** e o advogado **Arnessen Cintra** fazem esse acompanhamento de forma integrada.", { size: 21 }),
  heading(2, "Serviços incluídos"),
  cards([
    { step: "Contabilidade e Fiscal", title: "Empresa em dia com o fisco", tags: ["Metabase"], items: [
      "Apuração do Simples Nacional (PGDAS-D – Anexo II) e emissão do DAS. Controle do faturamento e dos limites do Simples.",
      "Obrigações junto à SEFAZ-PE e acompanhamento de regimes especiais de ICMS do setor de confecções.",
      "DEFIS e informe de rendimentos.",
      "Distribuição de lucros, pró-labore e orientação tributária.",
      "Certidões negativas e regularidade fiscal."] },
    { step: "Departamento Pessoal", title: "Folha correta todo mês", tags: ["Metabase"], items: [
      "Folha mensal com **cálculo da produção**, garantia do piso, DSR e horas extras conforme o modelo do planejamento.",
      "Admissões, experiências, férias, 13º, afastamentos e rescisões.",
      "eSocial (incluindo SST), FGTS Digital, DCTFWeb e guias.",
      "Conferência do **espelho de ponto** e do banco de horas, com alertas sobre jornada e horas extras.",
      "Controle de vencimentos de férias, experiências e exames periódicos.",
      "Obrigações anuais e informe de rendimentos dos empregados."] },
    { step: "Assessoria Jurídica Trabalhista", title: "Prevenção antes do problema", tags: ["Arnessen"], items: [
      "Consultoria preventiva ilimitada por WhatsApp, e-mail ou reunião: contratações, jornada, produção, advertências, suspensões e demissões.",
      "Revisão e atualização dos contratos, acordos de compensação/banco de horas, regulamento interno e termos individuais.",
      "Atendimento a fiscalizações do Ministério do Trabalho e defesa administrativa de autos de infração.",
      "Relacionamento com o sindicato e análise anual da nova Convenção Coletiva.",
      "Orientação nas rescisões de maior risco e acordos extrajudiciais (art. 855-B da CLT).",
      "1 reunião trimestral de revisão de riscos, junto com a contabilidade."] },
  ]),
  heading(3, "Investimento mensal"),
  prices(
    { dark: true, label: "Fase inicial · até 10 empregados", big: "1 salário mínimo", bigSize: 40, sub: "R$ 1.621,00 por mês em 2026, com contabilidade, fiscal, DP e jurídico incluídos.", split: ["Metabase: R$ 810,50", "Arnessen: R$ 810,50"] },
    { label: "Crescimento do quadro", big: "Reajuste negociado", bigSize: 30, sub: "O valor de 1 salário mínimo vale para até 10 empregados registrados. Se o quadro passar disso, o honorário terá reajuste negociado em comum acordo, proporcional ao volume de serviço." }),
  spacer(160),
  dataTable([4638, 3000, 2000], ["Serviço", "Responsável", "Valor mensal (R$)"], [
    ["Contabilidade, Fiscal e Departamento Pessoal (½ salário mínimo)", "Metabase Contabilidade", "810,50"],
    ["Assessoria Jurídica Trabalhista preventiva (½ salário mínimo)", "Arnessen Cintra – OAB/PE 70.236", "810,50"]],
    ["Total mensal (1 salário mínimo)", "", "1.621,00"], [2]),
  P("**Condições**", { size: 21, before: 200 }),
  bullet("**Vencimento:** dia 05 do mês subsequente ao da prestação do serviço."),
  bullet("**13ª parcela:** em dezembro, parcela adicional da parte contábil (R$ 810,50), para o 13º salário e a DEFIS."),
  bullet("**Taxa de implantação:** R$ 40,00 por admissão no lançamento inicial. Isenta se o planejamento (Proposta nº 001/2026) também for contratado."),
  bullet("**Reajuste:** anual e automático, acompanhando o salário mínimo nacional, independente de eventual renegociação por volume."),
  bullet("**Vigência:** prazo indeterminado, com 30 dias de aviso prévio para encerramento por qualquer das partes."),
  spacer(80),
  box("**Não incluídos (orçados à parte):** balanço patrimonial, balancetes e DRE, elaborados sob demanda com custo adicional; atuação em **reclamações trabalhistas e processos judiciais**, cobrados por ação conforme a tabela da OAB-PE; recuperação de períodos anteriores à contratação; regularização de pendências fiscais antigas; e custos de terceiros, como médico do trabalho, PGR/PCMSO, sistema de ponto, certificado digital, taxas e emolumentos.", true),
  spacer(60),
  small(juridicoNota + " A proposta é apresentada em conjunto para facilitar a gestão da contratante."),
  closing(4, "Responsabilidades da contratante", [
    "Enviar até o dia 25 de cada mês as informações para a folha: produção por empregado, ponto, faltas, admissões e desligamentos.",
    "Enviar as notas fiscais do mês até o dia 5 do mês seguinte.",
    "Consultar a assessoria **antes** de admitir, demitir ou mudar jornada e forma de pagamento."
  ]),
]);

(async () => {
  fs.writeFileSync(path.join(__dirname, "01-proposta-planejamento-trabalhista.docx"), await Packer.toBuffer(p1));
  fs.writeFileSync(path.join(__dirname, "02-proposta-acompanhamento-mensal.docx"), await Packer.toBuffer(p2));
  console.log("ok");
})();
