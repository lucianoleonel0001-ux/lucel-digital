require('dotenv').config();
const fetch      = require('node-fetch');
const express    = require('express');
const multer     = require('multer');
const nodemailer = require('nodemailer');
const fs         = require('fs');
const path       = require('path');
const crypto     = require('crypto');
const mammoth    = require('mammoth');
const pdfParse   = require('pdf-parse');
const { Document, Packer, Paragraph, TextRun, Header, Footer,
        AlignmentType, PageNumber, PageBreak, BorderStyle, TabStopType } = require('docx');

const app = express();
app.use(express.json());
app.use(express.static('.'));

const ADMIN_KEY     = process.env.ADMIN_KEY         || 'lucel2026';
const EMAIL_USER    = process.env.EMAIL_USER        || '';
const EMAIL_PASS    = process.env.EMAIL_PASS        || '';
const WHATS_TOKEN   = process.env.WHATS_TOKEN       || '';
const WHATS_NUM     = process.env.WHATS_NUM         || '5511934964127';
const BASE_URL      = process.env.BASE_URL          || 'https://lucel-digital.onrender.com';
const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY || '';

const DB_FILE     = './pedidos.json';
const UPLOAD_DIR  = './uploads';
const ENTREGA_DIR = './entregas';
[UPLOAD_DIR, ENTREGA_DIR].forEach(d => { if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true }); });

function lerPedidos() {
  if (!fs.existsSync(DB_FILE)) return [];
  try { return JSON.parse(fs.readFileSync(DB_FILE, 'utf8')); } catch { return []; }
}
function salvarPedidos(p) { fs.writeFileSync(DB_FILE, JSON.stringify(p, null, 2)); }

const multerStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename:    (req, file, cb) => cb(null, (req.params.id || 'tmp') + '_' + Date.now() + path.extname(file.originalname))
});
const upload = multer({ storage: multerStorage, limits: { fileSize: 50 * 1024 * 1024 } });

const transporter = nodemailer.createTransport({ service: 'gmail', auth: { user: EMAIL_USER, pass: EMAIL_PASS } });
async function enviarEmail(para, assunto, html) {
  if (!EMAIL_USER) return;
  try { await transporter.sendMail({ from: '"Lucel Digital" <' + EMAIL_USER + '>', to: para, subject: assunto, html }); }
  catch (e) { console.error('[EMAIL]', e.message); }
}
async function enviarWhatsApp(numero, msg) {
  if (!WHATS_TOKEN) return;
  const num = numero.replace(/\D/g, '');
  try { await fetch('https://api.callmebot.com/whatsapp.php?phone=+' + num + '&text=' + encodeURIComponent(msg) + '&apikey=' + WHATS_TOKEN); }
  catch (e) { console.error('[WHATS]', e.message); }
}

const FORMATOS = {
  '14x21': { pageW:7938,  pageH:11906, mTop:992,  mBot:992,  mEsq:1134, mDir:1134, mCab:482, mRod:482, corpoSize:24, tituloSize:32, cabSize:14, rodSize:20, recuo:482 },
  '16x23': { pageW:9072,  pageH:13032, mTop:1134, mBot:1134, mEsq:1361, mDir:1361, mCab:567, mRod:567, corpoSize:24, tituloSize:32, cabSize:14, rodSize:20, recuo:567 },
  'A4':    { pageW:11906, pageH:16838, mTop:1418, mBot:1418, mEsq:1701, mDir:1701, mCab:709, mRod:709, corpoSize:24, tituloSize:32, cabSize:14, rodSize:20, recuo:709 }
};
function getFmt(formato) {
  if (!formato) return FORMATOS['14x21'];
  if (formato.includes('16')) return FORMATOS['16x23'];
  if (formato.includes('A4') || formato.includes('29')) return FORMATOS['A4'];
  return FORMATOS['14x21'];
}

// ── CHAMAR CLAUDE API ──
async function chamarClaude(prompt, maxTokens) {
  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': ANTHROPIC_KEY,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: maxTokens || 8000,
      messages: [{ role: 'user', content: prompt }]
    })
  });
  const data = await resp.json();
  if (data.error) throw new Error('API Claude: ' + data.error.message);
  if (!data.content || !data.content[0]) throw new Error('API Claude: resposta vazia - ' + JSON.stringify(data).substring(0, 200));
  return data.content[0].text;
}

// ══════════════════════════════════════════
// ETAPA 1 — LER O LIVRO COMPLETO
// Extrai todo o texto do arquivo preservando
// a estrutura com HTML para capturar negritos
// ══════════════════════════════════════════
async function etapa1_lerLivro(filePath) {
  console.log('[ETAPA1] Lendo arquivo:', path.basename(filePath));
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.pdf') {
    const buf = fs.readFileSync(filePath);
    const data = await pdfParse(buf);
    const texto = data.text || '';
    console.log('[ETAPA1] PDF - chars:', texto.length);
    return { texto, html: null };
  } else {
    // Para DOCX: extrair HTML para preservar negritos, e texto limpo para análise
    const resultHtml  = await mammoth.convertToHtml({ path: filePath });
    const resultTexto = await mammoth.extractRawText({ path: filePath });
    console.log('[ETAPA1] DOCX - chars:', resultTexto.value.length);
    return { texto: resultTexto.value, html: resultHtml.value };
  }
}

// ══════════════════════════════════════════
// ETAPA 2 — IDENTIFICAR A ESTRUTURA
// Claude lê o livro todo e identifica:
// título, autor, credencial, lista de seções
// ══════════════════════════════════════════
async function etapa2_identificarEstrutura(texto) {
  console.log('[ETAPA2] Identificando estrutura...');

  const prompt =
    'Leia este livro completo e retorne APENAS JSON valido (sem markdown):\n' +
    '{"titulo":"titulo exato do livro","autor":"nome completo do autor ou vazio","credencial":"ex: Pastor, Dr. ou vazio","secoes":["titulo exato da secao 1","titulo exato da secao 2"]}\n\n' +
    'INSTRUCOES:\n' +
    '- Use o titulo EXATO como aparece no texto\n' +
    '- Liste TODAS as secoes: prefacio, introducao, capitulos, conclusao, etc\n' +
    '- Mantenha a numeracao exata dos capitulos (ex: "CAPITULO 1 - NOME")\n\n' +
    'LIVRO:\n' + texto.substring(0, 45000);

  const resposta = await chamarClaude(prompt, 3000);
  const estrutura = JSON.parse(resposta.replace(/```json|```/g, '').trim());
  console.log('[ETAPA2] Titulo:', estrutura.titulo);
  console.log('[ETAPA2] Autor:', estrutura.autor);
  console.log('[ETAPA2] Secoes:', estrutura.secoes ? estrutura.secoes.length : 0);
  return estrutura;
}

// ══════════════════════════════════════════
// ETAPA 3 — EXTRAIR CONTEÚDO DE CADA SEÇÃO
// Para cada seção: extrai parágrafos preservando
// texto original, negritos e numeração
// ══════════════════════════════════════════
async function etapa3_extrairConteudo(texto, html, secoes) {
  console.log('[ETAPA3] Processando', secoes.length, 'secoes...');
  const capitulos = [];

  for (let i = 0; i < secoes.length; i++) {
    const titulo = secoes[i];
    const proximo = secoes[i + 1] || null;

    // Localizar o texto desta seção
    const busca = titulo.substring(0, 35);
    let inicio = texto.indexOf(busca);
    if (inicio < 0) {
      console.log('[ETAPA3] Secao', i+1, 'nao encontrada:', busca);
      capitulos.push({ numero: String(i+1), titulo: titulo, paragrafos: [] });
      continue;
    }

    // Pular o título, pegar só o conteúdo
    const fimTitulo = texto.indexOf('\n', inicio);
    const inicioConteudo = fimTitulo > 0 ? fimTitulo + 1 : inicio;

    let fim = texto.length;
    if (proximo) {
      const buscaProx = proximo.substring(0, 35);
      const idxProx = texto.indexOf(buscaProx, inicioConteudo + 10);
      if (idxProx > inicioConteudo) fim = idxProx;
    }

    const conteudo = texto.substring(inicioConteudo, Math.min(fim, inicioConteudo + 12000)).trim();
    console.log('[ETAPA3] Secao', i+1, '"' + titulo.substring(0,30) + '" -', conteudo.length, 'chars');

    if (!conteudo || conteudo.length < 5) {
      capitulos.push({ numero: String(i+1), titulo: titulo, paragrafos: [] });
      continue;
    }

    // Claude processa o conteúdo preservando tudo
    const prompt =
      'Processe o texto abaixo e retorne APENAS JSON valido:\n' +
      '{"numero":"' + (i+1) + '","titulo":"' + titulo.replace(/"/g, '\\"') + '","paragrafos":[{"tipo":"normal","texto":"texto exato","negrito":false}]}\n\n' +
      'REGRAS CRITICAS - SIGA EXATAMENTE:\n' +
      '1. NAO adicione nem remova NENHUMA palavra do texto original\n' +
      '2. Preserve EXATAMENTE os negritos - marque com **texto** inline onde ha negrito no original\n' +
      '3. Mantenha a numeracao de titulos e subtitulos exatamente como esta\n' +
      '4. Corrija quebras de linha indevidas (una linhas que sao continuacao do mesmo paragrafo)\n' +
      '5. tipo pode ser: "normal" para paragrafos ou "subtitulo" para subtitulos internos\n' +
      '6. Inclua TODOS os paragrafos sem pular nenhum\n\n' +
      'TEXTO DA SECAO:\n' + conteudo;

    try {
      const resposta = await chamarClaude(prompt, 8000);
      const clean = resposta.replace(/```json|```/g, '').trim();
      const cap = JSON.parse(clean);
      const paragrafos = Array.isArray(cap.paragrafos) ? cap.paragrafos : [];
      capitulos.push({
        numero: String(i+1),
        titulo: String(cap.titulo || titulo),
        paragrafos: paragrafos
      });
      console.log('[ETAPA3] Secao', i+1, 'OK -', paragrafos.length, 'paragrafos');
    } catch (e) {
      console.error('[ETAPA3] Erro secao', i+1, ':', e.message);
      // Fallback: dividir por linhas em branco
      const paras = conteudo.split(/\n{2,}/).filter(p => p.trim().length > 0).map(p => ({
        tipo: 'normal',
        texto: p.trim().replace(/\n/g, ' '),
        negrito: false
      }));
      capitulos.push({ numero: String(i+1), titulo: titulo, paragrafos: paras });
      console.log('[ETAPA3] Fallback -', paras.length, 'paragrafos');
    }
  }

  return capitulos;
}

// ══════════════════════════════════════════
// ETAPA 4 — DIAGRAMAR
// Gera o DOCX com todas as especificações
// ══════════════════════════════════════════
async function etapa4_diagramar(estrutura, pedido, outputPath) {
  console.log('[ETAPA4] Gerando DOCX...');
  const fmt    = getFmt(pedido.formato);
  const titulo = String(estrutura.titulo || pedido.titulo || 'SEM TITULO');
  const autor  = String(estrutura.autor || '');
  const cred   = String(estrutura.credencial || '');
  const caps   = estrutura.capitulos || [];
  const PRETO  = '000000';
  const FONTE  = 'Palatino Linotype';

  function linhaTenue(a, d) {
    return new Paragraph({ spacing:{before:a||0,after:d||0}, border:{bottom:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}}, children:[new TextRun('')] });
  }
  function linhaEspessa(a, d) {
    return new Paragraph({ spacing:{before:a||0,after:d||0}, border:{bottom:{style:BorderStyle.SINGLE,size:8,color:PRETO,space:1}}, children:[new TextRun('')] });
  }
  function br() { return new Paragraph({ children:[new PageBreak()], spacing:{before:0,after:0} }); }
  function vazio() { return new Paragraph({ spacing:{before:0,after:0}, children:[new TextRun('')] }); }

  function makeRuns(texto, boldForcado) {
    if (!texto) return [new TextRun({ text:'', font:FONTE, size:fmt.corpoSize, color:PRETO })];
    // Processar marcação **negrito** inline
    const parts = String(texto).split(/(\*\*[^*]+\*\*)/g);
    return parts.filter(Boolean).map(function(p) {
      const isBold = !!boldForcado || /^\*\*[^*]+\*\*$/.test(p);
      return new TextRun({ text: p.replace(/\*\*/g, ''), font:FONTE, size:fmt.corpoSize, bold:isBold, color:PRETO });
    });
  }

  function pCorpo(texto, negrito) {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing:{ before:0, after:0, line:276, lineRule:'auto' },
      indent:{ firstLine:fmt.recuo },
      children: makeRuns(texto, negrito)
    });
  }
  function pTituloSecao(texto) {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing:{ before:280, after:800 },
      children:[new TextRun({ text:String(texto).toUpperCase(), font:FONTE, size:fmt.tituloSize, bold:true, color:PRETO })]
    });
  }
  function pSubtitulo(texto) {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing:{ before:200, after:400 },
      children:[new TextRun({ text:String(texto), font:FONTE, size:fmt.tituloSize-4, bold:true, color:PRETO })]
    });
  }

  // Cabeçalho
  const cabecalho = new Header({ children:[new Paragraph({
    spacing:{before:0,after:0},
    tabStops:[{type:TabStopType.RIGHT, position:fmt.pageW-fmt.mEsq-fmt.mDir}],
    border:{bottom:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}},
    children:[
      new TextRun({ text:titulo.toUpperCase(), font:FONTE, size:fmt.cabSize, allCaps:true, color:PRETO }),
      new TextRun({ text:'\t' }),
      new TextRun({ text:autor, font:FONTE, size:fmt.cabSize, italics:true, color:PRETO })
    ]
  })]});

  // Rodapé
  const rodape = new Footer({ children:[new Paragraph({
    alignment:AlignmentType.CENTER,
    spacing:{before:0,after:0},
    border:{top:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}},
    children:[new TextRun({ children:[PageNumber.CURRENT], font:FONTE, size:fmt.rodSize, color:PRETO })]
  })]});

  // Página de rosto
  const rosto = [
    vazio(), vazio(), vazio(), vazio(),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:280}, children:[new TextRun({ text:titulo.toUpperCase(), font:FONTE, size:56, bold:true, color:PRETO })] }),
    linhaEspessa(0,280),
    linhaTenue(0,160),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:160}, children:[new TextRun({ text:'\u2014 \u2014', font:FONTE, size:28, color:PRETO })] }),
    linhaTenue(0,480),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:120}, children:[new TextRun({ text:autor.toUpperCase(), font:FONTE, size:44, bold:true, color:PRETO })] }),
    cred ? new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:0}, children:[new TextRun({ text:cred, font:FONTE, size:20, italics:true, color:PRETO })] }) : vazio(),
    br()
  ];

  // Sumário
  const sumario = [
    linhaTenue(0,200),
    new Paragraph({ alignment:AlignmentType.LEFT, spacing:{before:0,after:400}, children:[new TextRun({ text:'SUM\u00c1RIO', font:FONTE, size:fmt.tituloSize, bold:true, color:PRETO })] }),
    ...caps.map(function(cap) {
      return new Paragraph({
        alignment:AlignmentType.LEFT,
        spacing:{before:120,after:60},
        children:[new TextRun({ text:String(cap.numero) + '. ' + String(cap.titulo || ''), font:FONTE, size:24, color:PRETO })]
      });
    }),
    br()
  ];

  // Conteúdo
  const conteudo = [];
  caps.forEach(function(cap, i) {
    conteudo.push(linhaTenue(0,200));
    conteudo.push(pTituloSecao(String(cap.numero) + '. ' + String(cap.titulo || '')));
    var paras = Array.isArray(cap.paragrafos) ? cap.paragrafos : [];
    paras.forEach(function(p) {
      if (!p || !p.texto) return;
      if (p.tipo === 'subtitulo') conteudo.push(pSubtitulo(p.texto));
      else conteudo.push(pCorpo(p.texto, p.negrito));
    });
    if (i < caps.length - 1) conteudo.push(br());
  });

  const doc = new Document({
    styles:{ default:{ document:{ run:{ font:FONTE, size:24, color:PRETO } } } },
    sections:[{
      properties:{ page:{ size:{ width:fmt.pageW, height:fmt.pageH }, margin:{ top:fmt.mTop, bottom:fmt.mBot, left:fmt.mEsq, right:fmt.mDir, header:fmt.mCab, footer:fmt.mRod } } },
      headers:{ default:cabecalho },
      footers:{ default:rodape },
      children:[...rosto, ...sumario, ...conteudo]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buf);
  console.log('[ETAPA4] DOCX gerado com', caps.length, 'capitulos');
}

// ══════════════════════════════════════════
// ETAPA 5 — CONFERIR E REVISAR (pacote completo)
// Claude revisa ortografia e gramática
// ══════════════════════════════════════════
async function etapa5_revisar(capitulos) {
  console.log('[ETAPA5] Revisando ortografia e gramatica...');
  const capitulosRevisados = [];

  for (let i = 0; i < capitulos.length; i++) {
    const cap = capitulos[i];
    if (!cap.paragrafos || cap.paragrafos.length === 0) {
      capitulosRevisados.push(cap);
      continue;
    }

    // Juntar parágrafos para revisão
    const textoParaRevisar = cap.paragrafos.map(p => p.texto).join('\n\n');

    const prompt =
      'Revise o texto abaixo corrigindo APENAS erros de:\n' +
      '- Ortografia\n' +
      '- Concordancia verbal e nominal\n' +
      '- Pontuacao\n' +
      '- Regencia\n' +
      '- Uso correto de crase\n\n' +
      'REGRAS CRITICAS:\n' +
      '1. NAO mude o estilo nem o conteudo\n' +
      '2. NAO adicione nem remova paragrafos\n' +
      '3. Mantenha os negritos marcados com **texto**\n' +
      '4. Retorne APENAS o texto corrigido, paragrafo por paragrafo separados por linha em branco\n\n' +
      'TEXTO:\n' + textoParaRevisar;

    try {
      const resposta = await chamarClaude(prompt, 8000);
      const paragrafosRevisados = resposta.trim().split(/\n{2,}/).filter(p => p.trim().length > 0);

      const novosParas = cap.paragrafos.map(function(p, idx) {
        return {
          tipo: p.tipo,
          texto: paragrafosRevisados[idx] ? paragrafosRevisados[idx].trim() : p.texto,
          negrito: p.negrito
        };
      });

      capitulosRevisados.push({ numero: cap.numero, titulo: cap.titulo, paragrafos: novosParas });
      console.log('[ETAPA5] Secao', i+1, 'revisada');
    } catch (e) {
      console.error('[ETAPA5] Erro revisao secao', i+1, ':', e.message);
      capitulosRevisados.push(cap);
    }
  }

  return capitulosRevisados;
}

// ══════════════════════════════════════════
// FLUXO PRINCIPAL
// ══════════════════════════════════════════
async function executarDiagramacao(pedido) {
  console.log('=== INICIANDO:', pedido.id, '===');
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(function(p) { return p.id === pedido.id; });

  try {
    const arquivoPath = path.join(UPLOAD_DIR, pedido.arquivoOriginal);
    if (!fs.existsSync(arquivoPath)) throw new Error('Arquivo nao encontrado');

    // ETAPA 1: Ler o livro
    const { texto, html } = await etapa1_lerLivro(arquivoPath);
    if (!texto || texto.length < 50) throw new Error('Arquivo sem conteudo');

    // ETAPA 2: Identificar estrutura
    const estruturaBase = await etapa2_identificarEstrutura(texto);
    if (!estruturaBase.secoes || estruturaBase.secoes.length === 0) throw new Error('Nenhuma secao identificada');

    // ETAPA 3: Extrair conteúdo de cada seção
    const capitulos = await etapa3_extrairConteudo(texto, html, estruturaBase.secoes);

    // ETAPA 4 (opcional): Revisão ortográfica/gramatical
    let capitulosFinais = capitulos;
    console.log('[FLUXO] Executando revisao ortografica e gramatical...');
    capitulosFinais = await etapa5_revisar(capitulos);

    // ETAPA 5: Diagramar
    const estrutura = {
      titulo: estruturaBase.titulo || pedido.titulo,
      autor: estruturaBase.autor || '',
      credencial: estruturaBase.credencial || '',
      capitulos: capitulosFinais
    };

    const docxPath = path.join(ENTREGA_DIR, pedido.id + '_diagramado.docx');
    await etapa4_diagramar(estrutura, pedido, docxPath);

    // ETAPA 6: Entregar ao cliente
    const token = crypto.randomBytes(16).toString('hex');
    pedidos[idx].status        = 'pronto';
    pedidos[idx].arquivoDocx   = path.basename(docxPath);
    pedidos[idx].downloadToken = token;
    pedidos[idx].linkDownload  = BASE_URL + '/download/' + pedido.id + '/' + token;
    pedidos[idx].entreguEm     = new Date().toISOString();
    salvarPedidos(pedidos);

    const link = pedidos[idx].linkDownload;
    await enviarEmail(pedido.email, 'Seu livro diagramado esta pronto! — Lucel Digital',
      '<p>Ola, <strong>' + pedido.nome + '</strong>!</p>' +
      '<p>Seu livro <em>' + pedido.titulo + '</em> esta pronto!</p>' +
      '<p><a href="' + link + '" style="background:#c9a96e;color:#000;padding:12px 24px;text-decoration:none;font-weight:bold;display:inline-block">BAIXAR MEU LIVRO</a></p>' +
      '<p style="color:#666;font-size:13px">Link: ' + link + '</p>');
    await enviarWhatsApp(pedido.whats, pedido.nome + ', seu livro "' + pedido.titulo + '" esta pronto! Baixe aqui: ' + link + ' — Lucel Digital');
    console.log('=== CONCLUIDO:', pedido.id, '===');

  } catch (e) {
    console.error('=== ERRO:', e.message, '===');
    console.error(e.stack);
    if (idx !== -1) {
      pedidos[idx].status = 'erro';
      pedidos[idx].erro   = e.message;
      salvarPedidos(pedidos);
    }
  }
}

// ── ADMIN AUTH ──
function adminAuth(req, res, next) {
  if (req.headers['x-admin-key'] === ADMIN_KEY) return next();
  res.status(401).json({ erro: 'Nao autorizado' });
}

// ── ROTAS ──
app.post('/api/pedido', function(req, res) {
  const { nome, email, whats, titulo, pacote, formato, preco } = req.body;
  if (!nome || !email || !whats || !titulo) return res.status(400).json({ erro: 'Campos obrigatorios' });
  const id    = crypto.randomBytes(4).toString('hex').toUpperCase();
  const agora = new Date();
  const pedido = { id, nome, email, whats, titulo, pacote, formato, preco, status:'aguardando',
    data: agora.toLocaleDateString('pt-BR') + ' ' + agora.toLocaleTimeString('pt-BR', {hour:'2-digit',minute:'2-digit'}),
    criadoEm: agora.toISOString(), arquivoOriginal: null };
  const pedidos = lerPedidos();
  pedidos.unshift(pedido);
  salvarPedidos(pedidos);
  enviarWhatsApp(WHATS_NUM, 'NOVO PEDIDO #' + id + '\nCliente: ' + nome + '\nLivro: ' + titulo + '\nFormato: ' + formato + '\nValor: R$ ' + preco + ',00');
  res.json({ ok: true, id });
});

app.post('/api/pedido/:id/arquivo', upload.single('arquivo'), function(req, res) {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(function(p) { return p.id === req.params.id; });
  if (idx === -1) return res.status(404).json({ erro: 'Nao encontrado' });
  if (req.file) pedidos[idx].arquivoOriginal = req.file.filename;
  salvarPedidos(pedidos);
  res.json({ ok: true });
});

app.get('/api/arquivo/:id', adminAuth, function(req, res) {
  const p = lerPedidos().find(function(p) { return p.id === req.params.id; });
  if (!p || !p.arquivoOriginal) return res.status(404).send('Nao encontrado');
  res.download(path.join(UPLOAD_DIR, p.arquivoOriginal));
});

app.get('/download/:id/:token', function(req, res) {
  const p = lerPedidos().find(function(p) { return p.id === req.params.id && p.downloadToken === req.params.token; });
  if (!p) return res.status(404).send('Link invalido.');
  res.send('<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Download — Lucel Digital</title>' +
    '<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:sans-serif;background:#0d0d0d;color:#fff;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center}' +
    '.box{background:#1a1a1a;padding:48px 40px;max-width:480px;width:100%}h1{font-size:24px;color:#c9a96e;margin-bottom:8px}p{color:#888;font-size:14px;margin-bottom:32px}' +
    '.btn{display:block;background:#c9a96e;color:#0d0d0d;padding:14px 32px;font-weight:700;text-decoration:none;margin-bottom:12px}.btn:hover{background:#e0c08a}</style></head>' +
    '<body><div class="box"><h1>Seu livro esta pronto!</h1><p>' + p.titulo + '</p>' +
    '<a class="btn" href="/download-file/' + p.id + '/' + p.downloadToken + '/docx">Baixar DOCX</a>' +
    '<p style="margin-top:24px;font-size:12px;color:#555">Lucel Digital · CNPJ 37.871.182/0001-86</p></div></body></html>');
});

app.get('/download-file/:id/:token/:tipo', function(req, res) {
  const p = lerPedidos().find(function(p) { return p.id === req.params.id && p.downloadToken === req.params.token; });
  if (!p) return res.status(404).send('Link invalido.');
  if (req.params.tipo === 'docx' && p.arquivoDocx) return res.download(path.join(ENTREGA_DIR, p.arquivoDocx));
  res.status(404).send('Arquivo nao encontrado.');
});

app.get('/api/pedidos', adminAuth, function(req, res) { res.json(lerPedidos()); });

app.post('/api/pedido/:id/liberar', adminAuth, async function(req, res) {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(function(p) { return p.id === req.params.id; });
  if (idx === -1) return res.status(404).json({ erro: 'Nao encontrado' });
  pedidos[idx].status     = 'liberado';
  pedidos[idx].liberadoEm = new Date().toISOString();
  salvarPedidos(pedidos);
  const p = pedidos[idx];
  res.json({ ok: true });
  await enviarEmail(p.email, 'Seu servico foi iniciado — Lucel Digital',
    '<p>Ola, <strong>' + p.nome + '</strong>!</p><p>Seu pedido <strong>#' + p.id + '</strong> foi confirmado e a diagramacao do livro <em>' + p.titulo + '</em> foi iniciada.</p><p>Voce recebera os arquivos em ate 2 horas por e-mail e WhatsApp.</p>');
  await enviarWhatsApp(p.whats, 'Ola ' + p.nome + '! Pedido #' + p.id + ' confirmado. Diagramacao de "' + p.titulo + '" iniciada. Voce recebera em ate 2h. — Lucel Digital');
  executarDiagramacao(p).catch(function(e) { console.error('[BG ERRO]', e.message); });
});

app.post('/api/pedido/:id/cancelar', adminAuth, function(req, res) {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(function(p) { return p.id === req.params.id; });
  if (idx === -1) return res.status(404).json({ erro: 'Nao encontrado' });
  pedidos[idx].status = 'cancelado';
  salvarPedidos(pedidos);
  res.json({ ok: true });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('Lucel Digital rodando na porta ' + PORT);
  if (!ANTHROPIC_KEY) console.log('AVISO: ANTHROPIC_KEY nao configurada!');
});
