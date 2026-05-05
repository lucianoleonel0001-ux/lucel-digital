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

// ── CONFIG ──
const ADMIN_KEY     = process.env.ADMIN_KEY         || 'lucel2026';
const EMAIL_USER    = process.env.EMAIL_USER        || '';
const EMAIL_PASS    = process.env.EMAIL_PASS        || '';
const WHATS_TOKEN   = process.env.WHATS_TOKEN       || '';
const WHATS_NUM     = process.env.WHATS_NUM         || '5511934964127';
const BASE_URL      = process.env.BASE_URL          || 'https://lucel-digital.onrender.com';
const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY || '';

// ── STORAGE ──
const DB_FILE     = './pedidos.json';
const UPLOAD_DIR  = './uploads';
const ENTREGA_DIR = './entregas';
[UPLOAD_DIR, ENTREGA_DIR].forEach(d => { if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true }); });

function lerPedidos() {
  if (!fs.existsSync(DB_FILE)) return [];
  try { return JSON.parse(fs.readFileSync(DB_FILE, 'utf8')); } catch { return []; }
}
function salvarPedidos(p) { fs.writeFileSync(DB_FILE, JSON.stringify(p, null, 2)); }

// ── MULTER ──
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename:    (req, file, cb) => cb(null, (req.params.id || 'tmp') + '_' + Date.now() + path.extname(file.originalname))
});
const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 } });

// ── EMAIL ──
const transporter = nodemailer.createTransport({ service: 'gmail', auth: { user: EMAIL_USER, pass: EMAIL_PASS } });
async function enviarEmail(para, assunto, html) {
  if (!EMAIL_USER) return console.log('[EMAIL] não configurado');
  try { await transporter.sendMail({ from: `"Lucel Digital" <${EMAIL_USER}>`, to: para, subject: assunto, html }); }
  catch (e) { console.error('[EMAIL]', e.message); }
}

// ── WHATSAPP ──
async function enviarWhatsApp(numero, msg) {
  if (!WHATS_TOKEN) return console.log('[WHATS] não configurado');
  const num = numero.replace(/\D/g, '');
  try { await fetch(`https://api.callmebot.com/whatsapp.php?phone=+${num}&text=${encodeURIComponent(msg)}&apikey=${WHATS_TOKEN}`); }
  catch (e) { console.error('[WHATS]', e.message); }
}

// ── FORMATOS ──
const FORMATOS = {
  '14x21': { pageW:7938,  pageH:11906, mTop:992,  mBot:992,  mEsq:1134, mDir:1134, mCab:482, mRod:482, corpoSize:24, tituloSize:32, cabSize:14, rodSize:20, recuo:482 },
  '16x23': { pageW:9072,  pageH:13032, mTop:1134, mBot:1134, mEsq:1361, mDir:1361, mCab:567, mRod:567, corpoSize:24, tituloSize:32, cabSize:14, rodSize:20, recuo:567 },
  'A4':    { pageW:11906, pageH:16838, mTop:1418, mBot:1418, mEsq:1701, mDir:1701, mCab:709, mRod:709, corpoSize:24, tituloSize:32, cabSize:14, rodSize:20, recuo:709 }
};

function getFmt(formato) {
  if (!formato) return FORMATOS['14x21'];
  if (formato.includes('16')) return FORMATOS['16x23'];
  if (formato.includes('A4') || formato.includes('21×29')) return FORMATOS['A4'];
  return FORMATOS['14x21'];
}

// ── EXTRAIR TEXTO ──
async function extrairTexto(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  try {
    if (ext === '.pdf') {
      const buf = fs.readFileSync(filePath);
      const data = await pdfParse(buf);
      return data.text || '';
    } else if (ext === '.docx' || ext === '.doc') {
      const result = await mammoth.extractRawText({ path: filePath });
      return result.value || '';
    }
  } catch (e) {
    console.error('[EXTRAIR]', e.message);
    return '';
  }
  return '';
}

// ── ANALISAR ESTRUTURA COM CLAUDE ──
async function analisarEstrutura(texto, pedido) {
  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': ANTHROPIC_KEY,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 8000,
      messages: [{
        role: 'user',
        content: `Você é um sistema de diagramação editorial brasileiro. Analise o texto abaixo e retorne APENAS um JSON válido (sem markdown, sem explicações) com esta estrutura:

{
  "titulo": "título principal do livro",
  "subtitulo": "subtítulo se houver, senão string vazia",
  "autor": "nome do autor se encontrado, senão string vazia",
  "credencial": "credencial do autor se houver, senão string vazia",
  "capitulos": [
    {
      "numero": "1",
      "titulo": "TÍTULO DO CAPÍTULO",
      "paragrafos": [
        { "tipo": "normal", "texto": "texto exato do parágrafo", "negrito": false }
      ]
    }
  ]
}

REGRAS CRÍTICAS:
1. NÃO acrescente nem remova NENHUMA palavra do texto original
2. Preserve EXATAMENTE onde há negrito — marque com **texto** inline
3. Preserve numeração de títulos/subtítulos exatamente como estão
4. Corrija quebras de linha indevidas (una frases que continuam na próxima linha)
5. tipo pode ser: "normal", "subtitulo", "lista"
6. Inclua TODOS os parágrafos do texto, não pule nenhum
7. Se o texto tiver prefácio, introdução, etc, inclua como capítulos

TEXTO DO LIVRO:
${texto.substring(0, 25000)}`
      }]
    })
  });
  const data = await resp.json();
  const raw = data.content?.[0]?.text || '{}';
  const clean = raw.replace(/```json|```/g, '').trim();
  try { return JSON.parse(clean); }
  catch {
    console.error('[ESTRUTURA] parse error, raw:', raw.substring(0, 200));
    return { titulo: pedido.titulo, subtitulo:'', autor:'', credencial:'', capitulos:[] };
  }
}

// ── GERAR DOCX ──
async function gerarDocx(estrutura, pedido, outputPath) {
  const fmt    = getFmt(pedido.formato);
  const titulo = estrutura.titulo     || pedido.titulo || 'SEM TÍTULO';
  const subtit = estrutura.subtitulo  || '';
  const autor  = estrutura.autor      || '';
  const cred   = estrutura.credencial || '';
  const caps   = estrutura.capitulos  || [];
  const PRETO  = '000000';
  const FONTE  = 'Palatino Linotype';

  function linhaTenue(a=0, d=0) {
    return new Paragraph({ spacing:{before:a,after:d}, border:{bottom:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}}, children:[new TextRun('')] });
  }
  function linhaEspessa(a=0, d=0) {
    return new Paragraph({ spacing:{before:a,after:d}, border:{bottom:{style:BorderStyle.SINGLE,size:8,color:PRETO,space:1}}, children:[new TextRun('')] });
  }
  function br() { return new Paragraph({ children:[new PageBreak()], spacing:{before:0,after:0} }); }
  function vazio() { return new Paragraph({ spacing:{before:0,after:0}, children:[new TextRun('')] }); }

  function runs(texto, boldForcado=false) {
    const parts = texto.split(/(\*\*.*?\*\*)/g);
    return parts.filter(Boolean).map(p => {
      const isBold = boldForcado || /^\*\*.*\*\*$/.test(p);
      return new TextRun({ text: p.replace(/\*\*/g,''), font:FONTE, size:fmt.corpoSize, bold:isBold, color:PRETO });
    });
  }

  function pCorpo(texto, negrito=false) {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing:{ before:0, after:0, line:276, lineRule:'auto' },
      indent:{ firstLine:fmt.recuo },
      children: runs(texto, negrito)
    });
  }
  function pTitulo(texto) {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing:{ before:280, after:800 },
      children:[new TextRun({ text:texto.toUpperCase(), font:FONTE, size:fmt.tituloSize, bold:true, color:PRETO })]
    });
  }
  function pSubtitulo(texto) {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing:{ before:200, after:400 },
      children:[new TextRun({ text:texto, font:FONTE, size:fmt.tituloSize-4, bold:true, color:PRETO })]
    });
  }

  const cabecalho = new Header({ children:[
    new Paragraph({
      spacing:{before:0,after:0},
      tabStops:[{type:TabStopType.RIGHT, position:fmt.pageW-fmt.mEsq-fmt.mDir}],
      border:{bottom:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}},
      children:[
        new TextRun({ text:titulo.toUpperCase(), font:FONTE, size:fmt.cabSize, allCaps:true, color:PRETO }),
        new TextRun({ text:'\t' }),
        new TextRun({ text:autor, font:FONTE, size:fmt.cabSize, italics:true, color:PRETO })
      ]
    })
  ]});

  const rodape = new Footer({ children:[
    new Paragraph({
      alignment:AlignmentType.CENTER,
      spacing:{before:0,after:0},
      border:{top:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}},
      children:[new TextRun({ children:[PageNumber.CURRENT], font:FONTE, size:fmt.rodSize, color:PRETO })]
    })
  ]});

  const rosto = [
    vazio(), vazio(), vazio(), vazio(),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:280}, children:[new TextRun({ text:titulo.toUpperCase(), font:FONTE, size:56, bold:true, color:PRETO })] }),
    linhaEspessa(0,280),
    ...(subtit ? [new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:160}, children:[new TextRun({ text:subtit, font:FONTE, size:36, italics:true, color:PRETO })] })] : []),
    linhaTenue(0,160),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:160}, children:[new TextRun({ text:'— —', font:FONTE, size:28, color:PRETO })] }),
    linhaTenue(0,480),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:120}, children:[new TextRun({ text:autor.toUpperCase(), font:FONTE, size:44, bold:true, color:PRETO })] }),
    ...(cred ? [new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:0}, children:[new TextRun({ text:cred, font:FONTE, size:20, italics:true, color:PRETO })] })] : []),
    br()
  ];

  const sumario = [
    linhaTenue(0,200),
    new Paragraph({ alignment:AlignmentType.LEFT, spacing:{before:0,after:400}, children:[new TextRun({ text:'SUMÁRIO', font:FONTE, size:fmt.tituloSize, bold:true, color:PRETO })] }),
    ...caps.map(cap => new Paragraph({
      alignment:AlignmentType.LEFT,
      spacing:{before:120,after:60},
      children:[new TextRun({ text:(cap.numero ? cap.numero+'. ':'')+cap.titulo, font:FONTE, size:24, color:PRETO })]
    })),
    br()
  ];

  const conteudo = [];
  caps.forEach((cap, i) => {
    conteudo.push(linhaTenue(0,200));
    conteudo.push(pTitulo((cap.numero ? cap.numero+'. ':'')+cap.titulo));
    (cap.paragrafos||[]).forEach(p => {
      if (p.tipo === 'subtitulo') conteudo.push(pSubtitulo(p.texto));
      else conteudo.push(pCorpo(p.texto, p.negrito||false));
    });
    if (i < caps.length-1) conteudo.push(br());
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
  console.log('[DOCX] Gerado:', outputPath);
}

// ── DIAGRAMAÇÃO AUTOMÁTICA ──
async function executarDiagramacao(pedido) {
  console.log('[DIAG] Iniciando:', pedido.id);
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(p => p.id === pedido.id);
  try {
    const arquivoPath = path.join(UPLOAD_DIR, pedido.arquivoOriginal);
    if (!fs.existsSync(arquivoPath)) throw new Error('Arquivo não encontrado: ' + arquivoPath);

    console.log('[DIAG] Extraindo texto...');
    const texto = await extrairTexto(arquivoPath);
    console.log('[DIAG] Texto extraído:', texto.length, 'chars');
    if (!texto || texto.length < 30) throw new Error('Texto vazio ou muito curto');

    console.log('[DIAG] Analisando estrutura com Claude...');
    const estrutura = await analisarEstrutura(texto, pedido);
    console.log('[DIAG] Capítulos encontrados:', estrutura.capitulos?.length || 0);

    console.log('[DIAG] Gerando DOCX...');
    const docxPath = path.join(ENTREGA_DIR, pedido.id + '_diagramado.docx');
    await gerarDocx(estrutura, pedido, docxPath);

    const token = crypto.randomBytes(16).toString('hex');
    pedidos[idx].status        = 'pronto';
    pedidos[idx].arquivoDocx   = path.basename(docxPath);
    pedidos[idx].downloadToken = token;
    pedidos[idx].linkDownload  = `${BASE_URL}/download/${pedido.id}/${token}`;
    pedidos[idx].entreguEm     = new Date().toISOString();
    salvarPedidos(pedidos);

    const link = pedidos[idx].linkDownload;
    await enviarEmail(pedido.email, '🎉 Seu livro diagramado está pronto! — Lucel Digital', `
      <p>Olá, <strong>${pedido.nome}</strong>!</p>
      <p>Seu livro <em>${pedido.titulo}</em> está pronto!</p>
      <p><a href="${link}" style="background:#c9a96e;color:#000;padding:12px 24px;text-decoration:none;font-weight:bold;display:inline-block">📥 BAIXAR MEU LIVRO</a></p>
      <p style="color:#666;font-size:13px">Link: ${link}</p>
      <br><p><strong>Lucel Digital</strong></p>
    `);
    await enviarWhatsApp(pedido.whats, `🎉 ${pedido.nome}, seu livro "${pedido.titulo}" está pronto! Baixe aqui: ${link} — Lucel Digital`);
    console.log('[DIAG] Concluído:', pedido.id);

  } catch (e) {
    console.error('[DIAG] Erro:', e.message);
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
  res.status(401).json({ erro: 'Não autorizado' });
}

// ── ROTAS CLIENTE ──
app.post('/api/pedido', (req, res) => {
  const { nome, email, whats, titulo, pacote, formato, preco } = req.body;
  if (!nome || !email || !whats || !titulo) return res.status(400).json({ erro: 'Campos obrigatórios' });
  const id    = crypto.randomBytes(4).toString('hex').toUpperCase();
  const agora = new Date();
  const pedido = { id, nome, email, whats, titulo, pacote, formato, preco, status:'aguardando',
    data: agora.toLocaleDateString('pt-BR')+' '+agora.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}),
    criadoEm: agora.toISOString(), arquivoOriginal: null };
  const pedidos = lerPedidos();
  pedidos.unshift(pedido);
  salvarPedidos(pedidos);
  enviarWhatsApp(WHATS_NUM, `🔔 NOVO PEDIDO #${id}\nCliente: ${nome}\nLivro: ${titulo}\nFormato: ${formato}\nValor: R$ ${preco},00\nWhatsApp: ${whats}`);
  res.json({ ok: true, id });
});

app.post('/api/pedido/:id/arquivo', upload.single('arquivo'), (req, res) => {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(p => p.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  if (req.file) pedidos[idx].arquivoOriginal = req.file.filename;
  salvarPedidos(pedidos);
  res.json({ ok: true });
});

app.get('/api/arquivo/:id', adminAuth, (req, res) => {
  const p = lerPedidos().find(p => p.id === req.params.id);
  if (!p?.arquivoOriginal) return res.status(404).send('Não encontrado');
  res.download(path.join(UPLOAD_DIR, p.arquivoOriginal));
});

app.get('/download/:id/:token', (req, res) => {
  const p = lerPedidos().find(p => p.id === req.params.id && p.downloadToken === req.params.token);
  if (!p) return res.status(404).send('Link inválido ou expirado.');
  res.send(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Download — Lucel Digital</title>
<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:sans-serif;background:#0d0d0d;color:#fff;display:flex;align-items:center;justify-content:center;min-height:100vh;text-align:center}
.box{background:#1a1a1a;padding:48px 40px;max-width:480px;width:100%}h1{font-size:24px;color:#c9a96e;margin-bottom:8px}p{color:#888;font-size:14px;margin-bottom:32px}
.btn{display:block;background:#c9a96e;color:#0d0d0d;padding:14px 32px;font-weight:700;text-decoration:none;margin-bottom:12px;font-size:15px}.btn:hover{background:#e0c08a}</style></head>
<body><div class="box"><h1>Seu livro está pronto!</h1><p>${p.titulo}</p>
<a class="btn" href="/download-file/${p.id}/${p.downloadToken}/docx">📝 Baixar DOCX</a>
<p style="margin-top:24px;font-size:12px;color:#555">Lucel Digital · CNPJ 37.871.182/0001-86</p>
</div></body></html>`);
});

app.get('/download-file/:id/:token/:tipo', (req, res) => {
  const p = lerPedidos().find(p => p.id === req.params.id && p.downloadToken === req.params.token);
  if (!p) return res.status(404).send('Link inválido.');
  if (req.params.tipo === 'docx' && p.arquivoDocx) return res.download(path.join(ENTREGA_DIR, p.arquivoDocx));
  res.status(404).send('Arquivo não encontrado.');
});

// ── ROTAS ADMIN ──
app.get('/api/pedidos', adminAuth, (req, res) => res.json(lerPedidos()));

app.post('/api/pedido/:id/liberar', adminAuth, async (req, res) => {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(p => p.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  pedidos[idx].status     = 'liberado';
  pedidos[idx].liberadoEm = new Date().toISOString();
  salvarPedidos(pedidos);
  const p = pedidos[idx];
  res.json({ ok: true });

  await enviarEmail(p.email, '✅ Seu serviço foi iniciado — Lucel Digital', `
    <p>Olá, <strong>${p.nome}</strong>!</p>
    <p>Seu pedido <strong>#${p.id}</strong> foi confirmado e a diagramação do livro <em>${p.titulo}</em> foi iniciada.</p>
    <p>Você receberá o arquivo em até 2 horas por e-mail e WhatsApp.</p>
    <br><p><strong>Lucel Digital</strong></p>
  `);
  await enviarWhatsApp(p.whats, `✅ Olá ${p.nome}! Pedido #${p.id} confirmado. Diagramação iniciada. Você receberá em até 2h. — Lucel Digital`);

  executarDiagramacao(p).catch(e => console.error('[DIAG BG]', e.message));
});

app.post('/api/pedido/:id/cancelar', adminAuth, (req, res) => {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(p => p.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  pedidos[idx].status = 'cancelado';
  salvarPedidos(pedidos);
  res.json({ ok: true });
});

app.get('/api/diag/:id', adminAuth, async (req, res) => {
  const p = lerPedidos().find(p => p.id === req.params.id);
  if (!p) return res.status(404).json({ erro: 'Não encontrado' });
  const info = { pedido: p, anthropicKey: ANTHROPIC_KEY ? 'configurada' : 'NAO' };
  if (p.arquivoOriginal) {
    const fp = path.join(UPLOAD_DIR, p.arquivoOriginal);
    info.arquivoExiste = fs.existsSync(fp);
    if (info.arquivoExiste) {
      info.arquivoTamanho = fs.statSync(fp).size;
      try { const t = await extrairTexto(fp); info.textoTamanho = t.length; info.textoPreview = t.substring(0,300); }
      catch(e) { info.erroExtracao = e.message; }
    }
  }
  res.json(info);
});

// ── START ──
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Lucel Digital rodando na porta ${PORT}`);
  if (!EMAIL_USER)    console.log('EMAIL_USER nao configurado');
  if (!WHATS_TOKEN)   console.log('WHATS_TOKEN nao configurado');
  if (!ANTHROPIC_KEY) console.log('ANTHROPIC_API_KEY nao configurada');
});
