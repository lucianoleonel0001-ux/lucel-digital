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

async function extrairTexto(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  console.log('[EXTRAIR] Arquivo:', path.basename(filePath), 'Ext:', ext);
  try {
    if (ext === '.pdf') {
      const buf = fs.readFileSync(filePath);
      const data = await pdfParse(buf);
      console.log('[EXTRAIR] PDF chars:', data.text ? data.text.length : 0);
      return data.text || '';
    } else {
      const result = await mammoth.extractRawText({ path: filePath });
      console.log('[EXTRAIR] DOCX chars:', result.value ? result.value.length : 0);
      return result.value || '';
    }
  } catch (e) {
    console.error('[EXTRAIR] Erro:', e.message);
    throw new Error('Falha ao extrair texto: ' + e.message);
  }
}

function detectarSecoes(texto) {
  const linhas = texto.split('\n');
  const secoes = [];
  const vistos = new Set();

  for (let i = 0; i < linhas.length; i++) {
    const linha = linhas[i].trim();
    if (!linha || linha.length > 100 || linha.length < 3) continue;
    if (vistos.has(linha)) continue;

    const ehTitulo =
      /^(PREF[AÁ]CIO|INTRODU[CÇ][AÃ]O|CONCLUS[AÃ]O|DEDICAT[OÓ]RIA|AGRADECIMENTOS?|APRESENTA[CÇ][AÃ]O|SUM[AÁ]RIO)/i.test(linha) ||
      /^CAP[IÍ]TULO\s+[\dIVXLC]+/i.test(linha) ||
      /^\d+\s*[-–—]\s+[A-ZÁÉÍÓÚÀÂÊÔÃÕÜ]/i.test(linha) ||
      /^\d+\.\s+[A-ZÁÉÍÓÚÀÂÊÔÃÕÜ]/i.test(linha);

    if (ehTitulo) {
      const linhaAntes = i > 0 ? linhas[i - 1].trim() : '';
      const linhaDepois = i < linhas.length - 1 ? linhas[i + 1].trim() : '';
      if (linhaAntes === '' || linhaDepois === '' || secoes.length === 0) {
        secoes.push(linha);
        vistos.add(linha);
      }
    }
  }
  return secoes;
}

async function analisarEstrutura(texto, pedido) {
  const secoes = detectarSecoes(texto);
  console.log('[DIAG] Secoes detectadas:', secoes.length, secoes.slice(0, 5));

  const linhasInicio = texto.split('\n').map(l => l.trim()).filter(l => l.length > 2 && l.length < 100);
  let tituloLivro = linhasInicio[0] || pedido.titulo;

  let autor = '';
  const finalTexto = texto.substring(Math.max(0, texto.length - 2000));
  const matchAutor = finalTexto.match(/(Pr\.?\s+[A-Z][^\n]{5,40}|Pastor\s+[A-Z][^\n]{5,40}|Dr\.?\s+[A-Z][^\n]{5,40})/);
  if (matchAutor) autor = matchAutor[1].trim();

  if (secoes.length === 0) {
    console.log('[DIAG] Nenhuma secao detectada, usando Claude para identificar');
    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': ANTHROPIC_KEY, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 2000,
        messages: [{ role: 'user', content: 'Liste os titulos de todos capitulos e secoes deste livro. Retorne APENAS JSON: {"titulo":"titulo do livro","autor":"nome do autor","secoes":["titulo1","titulo2"]}\n\nTEXTO:\n' + texto.substring(0, 20000) }]
      })
    });
    const d = await resp.json();
    try {
      const parsed = JSON.parse((d.content[0].text || '{}').replace(/```json|```/g, '').trim());
      if (parsed.titulo) tituloLivro = parsed.titulo;
      if (parsed.autor) autor = parsed.autor;
      secoes.push(...(parsed.secoes || []));
    } catch(e) { console.error('[DIAG] Parse erro etapa1:', e.message); }
  }

  const capitulos = [];
  for (let i = 0; i < secoes.length; i++) {
    const tituloSec = secoes[i];
    const busca = tituloSec.substring(0, 40);
    let idxInicio = texto.indexOf(busca);
    if (idxInicio < 0) { capitulos.push({ numero: String(i+1), titulo: tituloSec, paragrafos: [] }); continue; }

    let idxFim = texto.length;
    if (secoes[i+1]) {
      const buscaProx = secoes[i+1].substring(0, 40);
      // Buscar a próxima seção a partir do fim do título atual
      const idx = texto.indexOf(buscaProx, idxInicio + busca.length + 1);
      if (idx > idxInicio) idxFim = idx;
    }

    // Pegar o texto da seção (sem o título)
    const fimTitulo = texto.indexOf(String.fromCharCode(10), idxInicio);
    const inicioConteudo = fimTitulo > idxInicio ? fimTitulo + 1 : idxInicio;
    const textoSec = texto.substring(inicioConteudo, Math.min(idxFim, inicioConteudo + 15000));
    console.log('[DIAG] Secao', i+1, '- chars:', textoSec.length, '- titulo:', tituloSec.substring(0,30));

    try {
      const resp2 = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-api-key': ANTHROPIC_KEY, 'anthropic-version': '2023-06-01' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 8000,
          messages: [{ role: 'user', content: 'Processe este trecho e retorne APENAS JSON valido:\n{"numero":"' + (i+1) + '","titulo":"' + tituloSec.replace(/"/g, '\\"') + '","paragrafos":[{"tipo":"normal","texto":"texto exato do paragrafo","negrito":false}]}\n\nREGRAS:\n1. NAO adicione nem remova palavras\n2. Una linhas que sao continuacao do mesmo paragrafo\n3. tipo: normal ou subtitulo\n4. Inclua TODOS os paragrafos sem pular nenhum\n\nTEXTO:\n' + textoSec }]
        })
      });
      const d2 = await resp2.json();
      const raw2 = (d2.content[0].text || '{}').replace(/```json|```/g, '').trim();
      const cap = JSON.parse(raw2);
      capitulos.push({
        numero: String(cap.numero || i+1),
        titulo: String(cap.titulo || tituloSec),
        paragrafos: Array.isArray(cap.paragrafos) ? cap.paragrafos : []
      });
      console.log('[DIAG] Secao', i+1, '/', secoes.length, '- paragrafos:', cap.paragrafos ? cap.paragrafos.length : 0);
    } catch(e) {
      console.error('[DIAG] Erro secao', i+1, ':', e.message);
      capitulos.push({ numero: String(i+1), titulo: tituloSec, paragrafos: [] });
    }
  }

  return { titulo: tituloLivro, subtitulo: '', autor, credencial: '', capitulos };
}

async function gerarDocx(estrutura, pedido, outputPath) {
  const fmt    = getFmt(pedido.formato);
  const titulo = String(estrutura.titulo || pedido.titulo || 'SEM TITULO');
  const subtit = String(estrutura.subtitulo || '');
  const autor  = String(estrutura.autor || '');
  const cred   = String(estrutura.credencial || '');
  const caps   = (estrutura.capitulos || []).map((cap, i) => ({
    numero: String(cap.numero || i+1),
    titulo: String(cap.titulo || 'Capitulo ' + (i+1)),
    paragrafos: Array.isArray(cap.paragrafos) ? cap.paragrafos : []
  }));
  const PRETO = '000000';
  const FONTE = 'Palatino Linotype';

  function linhaTenue(a, d) {
    return new Paragraph({ spacing:{before:a||0,after:d||0}, border:{bottom:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}}, children:[new TextRun('')] });
  }
  function linhaEspessa(a, d) {
    return new Paragraph({ spacing:{before:a||0,after:d||0}, border:{bottom:{style:BorderStyle.SINGLE,size:8,color:PRETO,space:1}}, children:[new TextRun('')] });
  }
  function br() { return new Paragraph({ children:[new PageBreak()], spacing:{before:0,after:0} }); }
  function vazio() { return new Paragraph({ spacing:{before:0,after:0}, children:[new TextRun('')] }); }

  function runs(texto, boldForcado) {
    if (!texto) return [new TextRun({ text:'', font:FONTE, size:fmt.corpoSize, color:PRETO })];
    const parts = String(texto).split(/(\*\*[^*]+\*\*)/g);
    return parts.filter(Boolean).map(function(p) {
      const isBold = boldForcado || /^\*\*[^*]+\*\*$/.test(p);
      return new TextRun({ text: p.replace(/\*\*/g, ''), font:FONTE, size:fmt.corpoSize, bold:!!isBold, color:PRETO });
    });
  }

  function pCorpo(texto, negrito) {
    return new Paragraph({ alignment:AlignmentType.JUSTIFIED, spacing:{before:0,after:0,line:276,lineRule:'auto'}, indent:{firstLine:fmt.recuo}, children:runs(texto, negrito) });
  }
  function pTitulo(texto) {
    return new Paragraph({ alignment:AlignmentType.LEFT, spacing:{before:280,after:800}, children:[new TextRun({ text:String(texto).toUpperCase(), font:FONTE, size:fmt.tituloSize, bold:true, color:PRETO })] });
  }
  function pSubtitulo(texto) {
    return new Paragraph({ alignment:AlignmentType.LEFT, spacing:{before:200,after:400}, children:[new TextRun({ text:String(texto), font:FONTE, size:fmt.tituloSize-4, bold:true, color:PRETO })] });
  }

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

  const rodape = new Footer({ children:[new Paragraph({
    alignment:AlignmentType.CENTER,
    spacing:{before:0,after:0},
    border:{top:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}},
    children:[new TextRun({ children:[PageNumber.CURRENT], font:FONTE, size:fmt.rodSize, color:PRETO })]
  })]});

  const rosto = [
    vazio(), vazio(), vazio(), vazio(),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:280}, children:[new TextRun({ text:titulo.toUpperCase(), font:FONTE, size:56, bold:true, color:PRETO })] }),
    linhaEspessa(0, 280),
    subtit ? new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:160}, children:[new TextRun({ text:subtit, font:FONTE, size:36, italics:true, color:PRETO })] }) : vazio(),
    linhaTenue(0, 160),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:160}, children:[new TextRun({ text:'\u2014 \u2014', font:FONTE, size:28, color:PRETO })] }),
    linhaTenue(0, 480),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:120}, children:[new TextRun({ text:autor.toUpperCase(), font:FONTE, size:44, bold:true, color:PRETO })] }),
    cred ? new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:0}, children:[new TextRun({ text:cred, font:FONTE, size:20, italics:true, color:PRETO })] }) : vazio(),
    br()
  ];

  const sumario = [
    linhaTenue(0, 200),
    new Paragraph({ alignment:AlignmentType.LEFT, spacing:{before:0,after:400}, children:[new TextRun({ text:'SUM\u00c1RIO', font:FONTE, size:fmt.tituloSize, bold:true, color:PRETO })] }),
    ...caps.map(function(cap) {
      return new Paragraph({ alignment:AlignmentType.LEFT, spacing:{before:120,after:60}, children:[new TextRun({ text:cap.numero + '. ' + cap.titulo, font:FONTE, size:24, color:PRETO })] });
    }),
    br()
  ];

  const conteudo = [];
  caps.forEach(function(cap, i) {
    conteudo.push(linhaTenue(0, 200));
    conteudo.push(pTitulo(cap.numero + '. ' + cap.titulo));
    cap.paragrafos.forEach(function(p) {
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
  console.log('[DOCX] Gerado com', caps.length, 'capitulos');
}

async function executarDiagramacao(pedido) {
  console.log('=== [DIAG] INICIANDO DIAGRAMACAO ===');
  console.log('[DIAG] ID:', pedido.id);
  console.log('[DIAG] Titulo:', pedido.titulo);
  console.log('[DIAG] Formato:', pedido.formato);
  console.log('[DIAG] Arquivo:', pedido.arquivoOriginal);
  console.log('[DIAG] ANTHROPIC_KEY configurada:', !!ANTHROPIC_KEY);
  console.log('[DIAG] ANTHROPIC_KEY primeiros chars:', ANTHROPIC_KEY ? ANTHROPIC_KEY.substring(0,15) + '...' : 'VAZIA');

  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(function(p) { return p.id === pedido.id; });

  try {
    // PASSO 1: Verificar arquivo
    const arquivoPath = path.join(UPLOAD_DIR, pedido.arquivoOriginal);
    console.log('[DIAG] PASSO 1 - Verificando arquivo:', arquivoPath);
    if (!fs.existsSync(arquivoPath)) throw new Error('Arquivo nao encontrado: ' + arquivoPath);
    console.log('[DIAG] PASSO 1 - OK! Tamanho:', fs.statSync(arquivoPath).size, 'bytes');

    // PASSO 2: Extrair texto
    console.log('[DIAG] PASSO 2 - Extraindo texto...');
    const texto = await extrairTexto(arquivoPath);
    console.log('[DIAG] PASSO 2 - OK! Chars extraidos:', texto ? texto.length : 0);
    if (!texto || texto.length < 30) throw new Error('Texto vazio - chars: ' + (texto ? texto.length : 0));

    // PASSO 3: Analisar estrutura
    console.log('[DIAG] PASSO 3 - Analisando estrutura com Claude...');
    const estrutura = await analisarEstrutura(texto, pedido);
    console.log('[DIAG] PASSO 3 - OK! Capitulos:', estrutura.capitulos ? estrutura.capitulos.length : 0);
    console.log('[DIAG] PASSO 3 - Titulo:', estrutura.titulo);
    console.log('[DIAG] PASSO 3 - Autor:', estrutura.autor);

    // PASSO 4: Gerar DOCX
    console.log('[DIAG] PASSO 4 - Gerando DOCX...');
    const docxPath = path.join(ENTREGA_DIR, pedido.id + '_diagramado.docx');
    await gerarDocx(estrutura, pedido, docxPath);
    console.log('[DIAG] PASSO 4 - OK! DOCX gerado:', docxPath);

    // PASSO 5: Salvar e notificar
    console.log('[DIAG] PASSO 5 - Salvando e notificando cliente...');
    const token = crypto.randomBytes(16).toString('hex');
    pedidos[idx].status        = 'pronto';
    pedidos[idx].arquivoDocx   = path.basename(docxPath);
    pedidos[idx].downloadToken = token;
    pedidos[idx].linkDownload  = BASE_URL + '/download/' + pedido.id + '/' + token;
    pedidos[idx].entreguEm     = new Date().toISOString();
    salvarPedidos(pedidos);

    const link = pedidos[idx].linkDownload;
    console.log('[DIAG] PASSO 5 - Enviando email para:', pedido.email);
    await enviarEmail(pedido.email, 'Seu livro diagramado esta pronto! — Lucel Digital',
      '<p>Ola, <strong>' + pedido.nome + '</strong>!</p><p>Seu livro <em>' + pedido.titulo + '</em> esta pronto!</p><p><a href="' + link + '">BAIXAR MEU LIVRO</a></p>');
    console.log('[DIAG] PASSO 5 - Enviando WhatsApp para:', pedido.whats);
    await enviarWhatsApp(pedido.whats, pedido.nome + ', seu livro "' + pedido.titulo + '" esta pronto! Baixe aqui: ' + link + ' — Lucel Digital');
    console.log('=== [DIAG] CONCLUIDO COM SUCESSO:', pedido.id, '===');

  } catch (e) {
    console.error('=== [DIAG] ERRO FATAL ===');
    console.error('[DIAG] Mensagem:', e.message);
    console.error('[DIAG] Stack:', e.stack);
    if (idx !== -1) { 
      pedidos[idx].status = 'erro'; 
      pedidos[idx].erro = e.message + ' | ' + (e.stack || '').substring(0, 200);
      salvarPedidos(pedidos); 
    }
  }
}

function adminAuth(req, res, next) {
  if (req.headers['x-admin-key'] === ADMIN_KEY) return next();
  res.status(401).json({ erro: 'Nao autorizado' });
}

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
  enviarWhatsApp(WHATS_NUM, 'NOVO PEDIDO #' + id + '\nCliente: ' + nome + '\nLivro: ' + titulo + '\nFormato: ' + formato + '\nValor: R$ ' + preco + ',00\nWhatsApp: ' + whats);
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
  res.send('<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Download — Lucel Digital</title><style>body{font-family:sans-serif;background:#0d0d0d;color:#fff;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center}.box{background:#1a1a1a;padding:48px 40px;max-width:480px;width:100%}h1{font-size:24px;color:#c9a96e;margin-bottom:8px}p{color:#888;font-size:14px;margin-bottom:32px}.btn{display:block;background:#c9a96e;color:#0d0d0d;padding:14px 32px;font-weight:700;text-decoration:none;margin-bottom:12px}</style></head><body><div class="box"><h1>Seu livro esta pronto!</h1><p>' + p.titulo + '</p><a class="btn" href="/download-file/' + p.id + '/' + p.downloadToken + '/docx">Baixar DOCX</a></div></body></html>');
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
    '<p>Ola, <strong>' + p.nome + '</strong>! Seu pedido #' + p.id + ' foi confirmado. A diagramacao do livro "' + p.titulo + '" foi iniciada. Voce recebera em ate 2 horas.</p>');
  await enviarWhatsApp(p.whats, 'Ola ' + p.nome + '! Pedido #' + p.id + ' confirmado. Diagramacao do livro "' + p.titulo + '" iniciada. Voce recebera em ate 2h. — Lucel Digital');
  executarDiagramacao(p).catch(function(e) { console.error('[DIAG BG]', e.message); });
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
  if (!EMAIL_USER) console.log('EMAIL nao configurado');
  if (!WHATS_TOKEN) console.log('WHATS nao configurado');
  if (!ANTHROPIC_KEY) console.log('ANTHROPIC_KEY nao configurada');
});
