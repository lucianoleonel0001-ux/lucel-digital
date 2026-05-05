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

// ══════════════════════════════════════
// ESPECIFICAÇÕES DE FORMATO
// Baseado no livro de referência 14x21
// ══════════════════════════════════════
const FORMATOS = {
  '14x21': {
    pageW: 7937, pageH: 11905,
    mTop: 992, mBot: 992, mEsq: 1134, mDir: 1134,
    mCab: 482, mRod: 482,
    // Tamanhos em half-points (1pt = 2 half-pts)
    corpoSz: 24,      // 12pt
    tituloCapSz: 32,  // 16pt
    sumarioTitSz: 32, // 16pt
    sumarioItemSz: 24,// 12pt
    capaMainSz: 90,   // 45pt
    capaSubSz: 36,    // 18pt
    capaAutorSz: 44,  // 22pt
    capaCredSz: 20,   // 10pt
    cabSz: 14,        // 7pt
    rodSz: 20,        // 10pt
    recuo: 482        // 0.85cm
  },
  '16x23': {
    pageW: 9072, pageH: 13032,
    mTop: 1134, mBot: 1134, mEsq: 1304, mDir: 1304,
    mCab: 567, mRod: 567,
    corpoSz: 24, tituloCapSz: 32, sumarioTitSz: 32, sumarioItemSz: 24,
    capaMainSz: 90, capaSubSz: 36, capaAutorSz: 44, capaCredSz: 20,
    cabSz: 14, rodSz: 20, recuo: 567
  },
  'A4': {
    pageW: 11906, pageH: 16838,
    mTop: 1418, mBot: 1418, mEsq: 1701, mDir: 1701,
    mCab: 709, mRod: 709,
    corpoSz: 24, tituloCapSz: 32, sumarioTitSz: 32, sumarioItemSz: 24,
    capaMainSz: 90, capaSubSz: 36, capaAutorSz: 44, capaCredSz: 20,
    cabSz: 14, rodSz: 20, recuo: 709
  }
};

function getFmt(formato) {
  if (!formato) return FORMATOS['14x21'];
  if (formato.includes('16')) return FORMATOS['16x23'];
  if (formato.includes('A4') || formato.includes('29')) return FORMATOS['A4'];
  return FORMATOS['14x21'];
}

// ══════════════════════════════════════
// CHAMAR CLAUDE
// ══════════════════════════════════════
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
  if (data.error) throw new Error('API: ' + data.error.message);
  if (!data.content || !data.content[0]) throw new Error('API: sem resposta - ' + JSON.stringify(data).substring(0, 150));
  return data.content[0].text;
}

// ══════════════════════════════════════
// ETAPA 1 — LER O LIVRO COMPLETO
// Extrai texto e HTML (para negritos)
// ══════════════════════════════════════
async function etapa1_lerLivro(filePath) {
  console.log('[1] Lendo:', path.basename(filePath));
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.pdf') {
    const buf = fs.readFileSync(filePath);
    const data = await pdfParse(buf);
    console.log('[1] PDF chars:', data.text.length);
    return { texto: data.text || '', imagens: [], tabelas: [] };
  }

  // DOCX: extrair HTML preservando tabelas, negritos e imagens
  const resultHtml = await mammoth.convertToHtml({
    path: filePath,
    options: {
      includeDefaultStyleMap: true
    }
  });

  const html = resultHtml.value;

  // Extrair tabelas do HTML e converter para texto estruturado
  const tabelas = [];
  const htmlSemTabelas = html.replace(/<table[^>]*>([\s\S]*?)<\/table>/gi, function(match, conteudoTabela) {
    // Converter tabela HTML em texto formatado com |
    const linhas = [];
    const rows = conteudoTabela.match(/<tr[^>]*>([\s\S]*?)<\/tr>/gi) || [];
    rows.forEach(function(row) {
      const cells = row.match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi) || [];
      const celulas = cells.map(function(cell) {
        return cell.replace(/<[^>]+>/g, '').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&nbsp;/g,' ').trim();
      });
      if (celulas.length > 0) linhas.push(celulas.join(' | '));
    });
    const tabelaTexto = '[TABELA]\n' + linhas.join('\n') + '\n[/TABELA]';
    tabelas.push(tabelaTexto);
    return tabelaTexto;
  });

  // Converter HTML para texto preservando negritos e estrutura
  const texto = htmlSemTabelas
    .replace(/<strong>([\s\S]*?)<\/strong>/gi, '**$1**')
    .replace(/<b>([\s\S]*?)<\/b>/gi, '**$1**')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<br[^>]*>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  console.log('[1] DOCX chars:', texto.length, '| Tabelas:', tabelas.length);
  return { texto, imagens: [], tabelas };
}

// ══════════════════════════════════════
// ETAPA 2 — IDENTIFICAR ESTRUTURA
// ══════════════════════════════════════
async function etapa2e3_estruturarLivro(texto) {
  console.log('[2] Identificando estrutura...');

  // PASSO A: Identificar titulo, autor e secoes
  const promptA =
    'Leia este livro e retorne APENAS JSON valido:\n' +
    '{"titulo":"...","subtitulo":"...","autor":"...","credencial":"...","secoes":["titulo exato 1","titulo exato 2"]}\n\n' +
    'Liste TODAS as secoes: prefacio, introducao, capitulos, conclusao.\n\n' +
    'LIVRO:\n' + texto.substring(0, 60000);

  const respostaA = await chamarClaude(promptA, 3000);
  const base = JSON.parse(respostaA.replace(/```json|```/g, '').trim());
  console.log('[2] Titulo:', base.titulo, '| Secoes:', base.secoes ? base.secoes.length : 0);

  // PASSO B: Processar cada secao separadamente
  const capitulos = [];
  const secoes = base.secoes || [];

  for (let i = 0; i < secoes.length; i++) {
    const titulo = secoes[i];
    const proximo = secoes[i+1] || null;

    const busca = titulo.substring(0, 35);
    let inicio = texto.indexOf(busca);
    if (inicio < 0) { capitulos.push({ numero: String(i+1), titulo, paragrafos: [] }); continue; }

    const fimTitulo = texto.indexOf('\n', inicio);
    const inicioConteudo = fimTitulo > 0 ? fimTitulo + 1 : inicio;
    let fim = texto.length;
    if (proximo) {
      const idx = texto.indexOf(proximo.substring(0, 35), inicioConteudo + 10);
      if (idx > inicioConteudo) fim = idx;
    }

    const conteudo = texto.substring(inicioConteudo, Math.min(fim, inicioConteudo + 8000)).trim();
    if (!conteudo || conteudo.length < 5) { capitulos.push({ numero: String(i+1), titulo, paragrafos: [] }); continue; }

    const promptB =
      'Retorne APENAS JSON valido com os paragrafos desta secao:\n' +
      '{"numero":"' + (i+1) + '","titulo":"' + titulo.replace(/"/g, '\\"') + '","paragrafos":[{"tipo":"normal","texto":"texto exato","negrito":false}]}\n\n' +
      'REGRAS: nao adicione palavras, corrija quebras de linha, preserve negritos com **texto**.\n\n' +
      'TEXTO:\n' + conteudo;

    try {
      const respostaB = await chamarClaude(promptB, 4000);
      const cap = JSON.parse(respostaB.replace(/```json|```/g, '').trim());
      capitulos.push({ numero: String(i+1), titulo: cap.titulo || titulo, paragrafos: cap.paragrafos || [] });
      console.log('[2] Secao', i+1, 'OK -', cap.paragrafos ? cap.paragrafos.length : 0, 'paragrafos');
    } catch(e) {
      console.error('[2] Erro secao', i+1, e.message);
      const paras = conteudo.split(/\n{2,}/).filter(p=>p.trim()).map(p=>({tipo:'normal',texto:p.trim().replace(/\n/g,' '),negrito:false}));
      capitulos.push({ numero: String(i+1), titulo, paragrafos: paras });
    }
  }

  return { titulo: base.titulo || '', subtitulo: base.subtitulo || '', autor: base.autor || '', credencial: base.credencial || '', capitulos };
}

// ══════════════════════════════════════
// ETAPA 4 — REVISÃO ORTOGRÁFICA
// Corrige ortografia, gramática,
// pontuação, regência, crase
// sem alterar o conteúdo
// ══════════════════════════════════════
async function etapa4_revisar(capitulos) {
  console.log('[4] Revisando ortografia e gramatica...');
  const revisados = [];

  for (let i = 0; i < capitulos.length; i++) {
    const cap = capitulos[i];
    if (!cap.paragrafos || cap.paragrafos.length === 0) {
      revisados.push(cap); continue;
    }

    const textoParas = cap.paragrafos.map(p => p.texto || '').join('\n\n');
    const prompt =
      'Revise o texto corrigindo APENAS erros de ortografia, concordancia, pontuacao, regencia e crase.\n' +
      'NAO mude o estilo, conteudo, paragrafos nem negritos (**texto**).\n' +
      'Retorne APENAS o texto corrigido, paragrafos separados por linha em branco:\n\n' +
      textoParas;

    try {
      const resposta = await chamarClaude(prompt, 8000);
      const parasRevisados = resposta.trim().split(/\n{2,}/).filter(p => p.trim());
      const novosParagrafos = cap.paragrafos.map(function(p, idx) {
        return { tipo: p.tipo, texto: parasRevisados[idx] ? parasRevisados[idx].trim() : p.texto, negrito: p.negrito };
      });
      revisados.push({ numero: cap.numero, titulo: cap.titulo, paragrafos: novosParagrafos });
      console.log('[4] Secao', i+1, 'revisada');
    } catch (e) {
      console.error('[4] Erro revisao', i+1, ':', e.message);
      revisados.push(cap);
    }
  }
  return revisados;
}

// ══════════════════════════════════════
// ETAPA 5 — DIAGRAMAR
// Gera o DOCX com as especificações
// exatas do livro de referência
// ══════════════════════════════════════
async function etapa5_diagramar(estrutura, pedido, outputPath) {
  console.log('[5] Gerando DOCX...');
  const fmt   = getFmt(pedido.formato);
  const titulo = String(estrutura.titulo || pedido.titulo || 'SEM TITULO');
  const subtit = String(estrutura.subtitulo || '');
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
    if (!texto) return [new TextRun({ text:'', font:FONTE, size:fmt.corpoSz, color:PRETO })];
    const parts = String(texto).split(/(\*\*[^*]+\*\*)/g);
    return parts.filter(Boolean).map(function(p) {
      const isBold = !!boldForcado || /^\*\*[^*]+\*\*$/.test(p);
      return new TextRun({ text: p.replace(/\*\*/g, ''), font:FONTE, size:fmt.corpoSz, bold:isBold, color:PRETO });
    });
  }

  function pCorpo(texto, negrito) {
    // Verificar se é uma tabela
    if (String(texto).startsWith('[TABELA]')) {
      const linhas = String(texto).replace('[TABELA]','').replace('[/TABELA]','').trim().split('\n');
      return linhas.map(function(linha) {
        return new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing:{ before:0, after:60 },
          children:[new TextRun({ text: linha, font:FONTE, size:fmt.corpoSz, color:PRETO })]
        });
      });
    }
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
      children:[new TextRun({ text:String(texto).toUpperCase(), font:FONTE, size:fmt.tituloCapSz, bold:true, color:PRETO })]
    });
  }
  function pSubtitulo(texto) {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing:{ before:200, after:400 },
      children:[new TextRun({ text:String(texto), font:FONTE, size:fmt.tituloCapSz-4, bold:true, color:PRETO })]
    });
  }

  // CABEÇALHO: título esquerda | autor direita
  const cabecalho = new Header({ children:[new Paragraph({
    spacing:{before:0,after:0},
    tabStops:[{type:TabStopType.RIGHT, position:fmt.pageW-fmt.mEsq-fmt.mDir}],
    border:{bottom:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}},
    children:[
      new TextRun({ text:titulo.toUpperCase(), font:FONTE, size:fmt.cabSz, allCaps:true, color:PRETO }),
      new TextRun({ text:'\t' }),
      new TextRun({ text:autor, font:FONTE, size:fmt.cabSz, italics:true, color:PRETO })
    ]
  })]});

  // RODAPÉ: número de página centralizado
  const rodape = new Footer({ children:[new Paragraph({
    alignment:AlignmentType.CENTER,
    spacing:{before:0,after:0},
    border:{top:{style:BorderStyle.SINGLE,size:2,color:PRETO,space:1}},
    children:[new TextRun({ children:[PageNumber.CURRENT], font:FONTE, size:fmt.rodSz, color:PRETO })]
  })]});

  // PÁGINA DE ROSTO
  const rosto = [
    vazio(), vazio(), vazio(), vazio(),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:280},
      children:[new TextRun({ text:titulo.toUpperCase(), font:FONTE, size:fmt.capaMainSz, bold:true, color:PRETO })] }),
    linhaEspessa(0, 280),
    subtit
      ? new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:160},
          children:[new TextRun({ text:subtit, font:FONTE, size:fmt.capaSubSz, italics:true, color:PRETO })] })
      : vazio(),
    linhaTenue(0, 160),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:160},
      children:[new TextRun({ text:'\u2014 \u2014', font:FONTE, size:24, color:PRETO })] }),
    linhaTenue(0, 480),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:120},
      children:[new TextRun({ text:autor.toUpperCase(), font:FONTE, size:fmt.capaAutorSz, bold:true, color:PRETO })] }),
    cred
      ? new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:0},
          children:[new TextRun({ text:cred, font:FONTE, size:fmt.capaCredSz, italics:true, color:PRETO })] })
      : vazio(),
    br()
  ];

  // SUMÁRIO
  const sumario = [
    linhaTenue(0, 200),
    new Paragraph({ alignment:AlignmentType.LEFT, spacing:{before:0,after:400},
      children:[new TextRun({ text:'SUM\u00c1RIO', font:FONTE, size:fmt.sumarioTitSz, bold:true, color:PRETO })] }),
    ...caps.map(function(cap) {
      return new Paragraph({ alignment:AlignmentType.LEFT, spacing:{before:120,after:60},
        children:[new TextRun({ text:String(cap.numero) + '. ' + String(cap.titulo || ''), font:FONTE, size:fmt.sumarioItemSz, color:PRETO })] });
    }),
    br()
  ];

  // CONTEÚDO DOS CAPÍTULOS
  const conteudo = [];
  caps.forEach(function(cap, i) {
    conteudo.push(linhaTenue(0, 200));
    conteudo.push(pTituloSecao(String(cap.numero) + '. ' + String(cap.titulo || '')));
    var paras = Array.isArray(cap.paragrafos) ? cap.paragrafos : [];
    paras.forEach(function(p) {
      if (!p || !p.texto) return;
      if (p.tipo === 'subtitulo') {
        conteudo.push(pSubtitulo(p.texto));
      } else {
        var result = pCorpo(p.texto, p.negrito);
        if (Array.isArray(result)) {
          result.forEach(function(r) { conteudo.push(r); });
        } else {
          conteudo.push(result);
        }
      }
    });
    if (i < caps.length - 1) conteudo.push(br());
  });

  const doc = new Document({
    styles:{ default:{ document:{ run:{ font:FONTE, size:fmt.corpoSz, color:PRETO } } } },
    sections:[{
      properties:{ page:{
        size:{ width:fmt.pageW, height:fmt.pageH },
        margin:{ top:fmt.mTop, bottom:fmt.mBot, left:fmt.mEsq, right:fmt.mDir, header:fmt.mCab, footer:fmt.mRod }
      }},
      headers:{ default:cabecalho },
      footers:{ default:rodape },
      children:[...rosto, ...sumario, ...conteudo]
    }]
  });

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buf);
  console.log('[5] DOCX gerado:', caps.length, 'capitulos');
}

// ══════════════════════════════════════
// ETAPA 6 — ENTREGAR AO CLIENTE
// Salva, gera link e notifica por
// email e WhatsApp
// ══════════════════════════════════════
async function etapa6_entregar(pedido, docxPath, pedidos, idx) {
  console.log('[6] Entregando ao cliente...');
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
  await enviarWhatsApp(pedido.whats,
    pedido.nome + ', seu livro "' + pedido.titulo + '" esta pronto! Baixe aqui: ' + link + ' — Lucel Digital');
  console.log('[6] Cliente notificado:', pedido.email);
}

// ══════════════════════════════════════
// FLUXO PRINCIPAL
// ══════════════════════════════════════
async function executarDiagramacao(pedido) {
  console.log('=== INICIO:', pedido.id, pedido.titulo, '===');
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(function(p) { return p.id === pedido.id; });

  try {
    const arquivoPath = path.join(UPLOAD_DIR, pedido.arquivoOriginal);
    if (!fs.existsSync(arquivoPath)) throw new Error('Arquivo nao encontrado: ' + arquivoPath);

    // ETAPA 1: Ler o livro completo
    const livro = await etapa1_lerLivro(arquivoPath);
    const texto = livro.texto || livro;
    if (!texto || texto.length < 50) throw new Error('Arquivo sem conteudo de texto');

    // ETAPAS 2+3: Estruturar livro completo em uma chamada
    const estruturaBase = await etapa2e3_estruturarLivro(texto);
    if (!estruturaBase.capitulos || estruturaBase.capitulos.length === 0) {
      throw new Error('Nenhum capitulo identificado no livro');
    }

    // ETAPA 4: Revisão ortográfica e gramatical
    const capitulosRevisados = await etapa4_revisar(estruturaBase.capitulos);

    // ETAPA 5: Diagramar o DOCX
    const estrutura = {
      titulo:     estruturaBase.titulo     || pedido.titulo,
      subtitulo:  estruturaBase.subtitulo  || '',
      autor:      estruturaBase.autor      || '',
      credencial: estruturaBase.credencial || '',
      capitulos:  capitulosRevisados
    };
    const docxPath = path.join(ENTREGA_DIR, pedido.id + '_diagramado.docx');
    await etapa5_diagramar(estrutura, pedido, docxPath);

    // ETAPA 6: Entregar ao cliente
    await etapa6_entregar(pedido, docxPath, pedidos, idx);
    console.log('=== CONCLUIDO:', pedido.id, '===');

  } catch (e) {
    console.error('=== ERRO:', e.message);
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
    '<p>Ola, <strong>' + p.nome + '</strong>!</p>' +
    '<p>Seu pedido <strong>#' + p.id + '</strong> foi confirmado. A diagramacao do livro <em>' + p.titulo + '</em> foi iniciada.</p>' +
    '<p>Voce recebera os arquivos em ate 2 horas por e-mail e WhatsApp.</p>');
  await enviarWhatsApp(p.whats,
    'Ola ' + p.nome + '! Pedido #' + p.id + ' confirmado. Diagramacao de "' + p.titulo + '" iniciada. Voce recebera em ate 2h. — Lucel Digital');
  executarDiagramacao(p).catch(function(e) { console.error('[BG]', e.message); });
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
