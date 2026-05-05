const express = require('express');
const multer = require('multer');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
app.use(express.json());
app.use(express.static('.'));

// ── CONFIG ── (ajuste no Render via Environment Variables)
const ADMIN_KEY   = process.env.ADMIN_KEY   || 'lucel2026';
const EMAIL_USER  = process.env.EMAIL_USER  || '';   // seu gmail
const EMAIL_PASS  = process.env.EMAIL_PASS  || '';   // senha de app gmail
const WHATS_TOKEN = process.env.WHATS_TOKEN || '';   // token CallMeBot ou Z-API
const WHATS_NUM   = process.env.WHATS_NUM   || '5511934964127';
const BASE_URL    = process.env.BASE_URL    || 'https://lucialdigital.onrender.com';

// ── STORAGE ──
const DB_FILE = './pedidos.json';
const UPLOAD_DIR = './uploads';
const ENTREGA_DIR = './entregas';
[UPLOAD_DIR, ENTREGA_DIR].forEach(d => { if (!fs.existsSync(d)) fs.mkdirSync(d); });

function lerPedidos() {
  if (!fs.existsSync(DB_FILE)) return [];
  try { return JSON.parse(fs.readFileSync(DB_FILE, 'utf-8')); } catch { return []; }
}
function salvarPedidos(p) { fs.writeFileSync(DB_FILE, JSON.stringify(p, null, 2)); }

// ── MULTER ──
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = req.path.includes('entregar') ? ENTREGA_DIR : UPLOAD_DIR;
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, req.params.id + '_' + Date.now() + ext);
  }
});
const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 } });

// ── EMAIL ──
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: { user: EMAIL_USER, pass: EMAIL_PASS }
});

async function enviarEmail(para, assunto, html) {
  if (!EMAIL_USER) return console.log('[EMAIL] Não configurado:', assunto);
  try {
    await transporter.sendMail({ from: `"Lucel Digital" <${EMAIL_USER}>`, to: para, subject: assunto, html });
    console.log('[EMAIL] Enviado para:', para);
  } catch (e) { console.error('[EMAIL] Erro:', e.message); }
}

// ── WHATSAPP (CallMeBot API) ──
async function enviarWhatsApp(numero, mensagem) {
  if (!WHATS_TOKEN) return console.log('[WHATS] Não configurado:', mensagem);
  const num = numero.replace(/\D/g, '');
  const url = `https://api.callmebot.com/whatsapp.php?phone=+${num}&text=${encodeURIComponent(mensagem)}&apikey=${WHATS_TOKEN}`;
  try {
    const res = await fetch(url);
    console.log('[WHATS] Status:', res.status);
  } catch (e) { console.error('[WHATS] Erro:', e.message); }
}

// ── MIDDLEWARE ADMIN ──
function adminAuth(req, res, next) {
  if (req.headers['x-admin-key'] === ADMIN_KEY) return next();
  res.status(401).json({ erro: 'Não autorizado' });
}

// ── ROTAS CLIENTE ──

// Receber novo pedido
app.post('/api/pedido', (req, res) => {
  const { nome, email, whats, titulo, pacote, formato, preco } = req.body;
  if (!nome || !email || !whats || !titulo) return res.status(400).json({ erro: 'Campos obrigatórios' });

  const id = crypto.randomBytes(4).toString('hex').toUpperCase();
  const agora = new Date();
  const pedido = {
    id,
    nome, email, whats, titulo, pacote, formato, preco,
    status: 'aguardando',
    data: agora.toLocaleDateString('pt-BR') + ' ' + agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }),
    criadoEm: agora.toISOString(),
    arquivoOriginal: null,
    arquivoEntrega: null,
    linkDownload: null
  };

  const pedidos = lerPedidos();
  pedidos.unshift(pedido);
  salvarPedidos(pedidos);

  // Aviso para o admin via WhatsApp
  enviarWhatsApp(WHATS_NUM,
    `🔔 NOVO PEDIDO #${id}\nCliente: ${nome}\nLivro: ${titulo}\nPacote: ${pacote}\nFormato: ${formato}\nValor: R$ ${preco},00\nWhatsApp: ${whats}`
  );

  res.json({ ok: true, id });
});

// Upload do arquivo do cliente (chamado junto com o pedido)
app.post('/api/pedido/:id/arquivo', upload.single('arquivo'), (req, res) => {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(p => p.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Pedido não encontrado' });
  pedidos[idx].arquivoOriginal = req.file ? req.file.filename : null;
  salvarPedidos(pedidos);
  res.json({ ok: true });
});

// Download do arquivo do cliente (admin)
app.get('/api/arquivo/:id', adminAuth, (req, res) => {
  const pedidos = lerPedidos();
  const p = pedidos.find(p => p.id === req.params.id);
  if (!p || !p.arquivoOriginal) return res.status(404).send('Arquivo não encontrado');
  res.download(path.join(UPLOAD_DIR, p.arquivoOriginal));
});

// Download do arquivo entregue (cliente via link)
app.get('/download/:id/:token', (req, res) => {
  const pedidos = lerPedidos();
  const p = pedidos.find(p => p.id === req.params.id);
  if (!p || !p.arquivoEntrega || p.downloadToken !== req.params.token) {
    return res.status(404).send('Link inválido ou expirado.');
  }
  res.download(path.join(ENTREGA_DIR, p.arquivoEntrega));
});

// ── ROTAS ADMIN ──

// Listar pedidos
app.get('/api/pedidos', adminAuth, (req, res) => {
  res.json(lerPedidos());
});

// Liberar serviço
app.post('/api/pedido/:id/liberar', adminAuth, async (req, res) => {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(p => p.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });

  pedidos[idx].status = 'liberado';
  pedidos[idx].liberadoEm = new Date().toISOString();
  salvarPedidos(pedidos);

  const p = pedidos[idx];

  // Email para o cliente
  await enviarEmail(p.email, '✅ Seu serviço foi iniciado — Lucel Digital', `
    <p>Olá, <strong>${p.nome}</strong>!</p>
    <p>Seu pedido <strong>#${p.id}</strong> foi confirmado e a diagramação do seu livro <em>${p.titulo}</em> foi iniciada.</p>
    <p>Você receberá o arquivo (.docx e .pdf) em até 2 horas por e-mail e WhatsApp.</p>
    <br><p>Atenciosamente,<br><strong>Lucel Digital</strong></p>
  `);

  // WhatsApp para o cliente
  await enviarWhatsApp(p.whats,
    `✅ Olá ${p.nome}! Seu pedido #${p.id} foi confirmado e a diagramação do livro "${p.titulo}" foi iniciada. Você receberá o arquivo em até 2 horas. — Lucel Digital`
  );

  res.json({ ok: true });
});

// Entregar arquivo diagramado
app.post('/api/pedido/:id/entregar', adminAuth, upload.single('arquivo'), async (req, res) => {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(p => p.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  if (!req.file) return res.status(400).json({ erro: 'Arquivo não enviado' });

  const token = crypto.randomBytes(16).toString('hex');
  pedidos[idx].status = 'pronto';
  pedidos[idx].arquivoEntrega = req.file.filename;
  pedidos[idx].downloadToken = token;
  pedidos[idx].linkDownload = `${BASE_URL}/download/${pedidos[idx].id}/${token}`;
  pedidos[idx].entreguEm = new Date().toISOString();
  salvarPedidos(pedidos);

  const p = pedidos[idx];
  const link = p.linkDownload;

  // Email para o cliente
  await enviarEmail(p.email, '🎉 Seu livro diagramado está pronto! — Lucel Digital', `
    <p>Olá, <strong>${p.nome}</strong>!</p>
    <p>O seu livro <em>${p.titulo}</em> foi diagramado com sucesso!</p>
    <p><strong>Clique no link abaixo para baixar seus arquivos (.docx e .pdf):</strong></p>
    <p><a href="${link}" style="background:#c9a96e;color:#000;padding:12px 24px;text-decoration:none;font-weight:bold;display:inline-block">📥 BAIXAR MEU LIVRO</a></p>
    <p style="color:#666;font-size:13px">Se o botão não funcionar, copie e cole este link: ${link}</p>
    <br><p>Obrigado pela confiança!<br><strong>Lucel Digital</strong></p>
  `);

  // WhatsApp para o cliente
  await enviarWhatsApp(p.whats,
    `🎉 ${p.nome}, seu livro "${p.titulo}" está pronto! Acesse o link para baixar: ${link} — Lucel Digital`
  );

  res.json({ ok: true, link });
});

// Cancelar pedido
app.post('/api/pedido/:id/cancelar', adminAuth, (req, res) => {
  const pedidos = lerPedidos();
  const idx = pedidos.findIndex(p => p.id === req.params.id);
  if (idx === -1) return res.status(404).json({ erro: 'Não encontrado' });
  pedidos[idx].status = 'cancelado';
  salvarPedidos(pedidos);
  res.json({ ok: true });
});

// ── START ──
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Lucel Digital rodando na porta ${PORT}`));
