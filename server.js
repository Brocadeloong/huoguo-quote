const http = require('http');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const PORT = process.env.PORT || 3000;
const LOG_FILE = path.join(__dirname, 'quote-log.jsonl');
const EXCEL_DIR = path.join(__dirname, 'excel');
const MAX_BODY = 1024 * 1024; // 1MB
const MIN_TOTAL = 100;

function send(res, status, payload) {
  res.writeHead(status, {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  });
  res.end(JSON.stringify(payload));
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => {
      data += chunk;
      if (data.length > MAX_BODY) {
        reject(new Error('Payload too large'));
        req.destroy();
      }
    });
    req.on('end', () => {
      try {
        resolve(JSON.parse(data || '{}'));
      } catch (err) {
        reject(err);
      }
    });
    req.on('error', reject);
  });
}

async function handleQuote(req, res) {
  let body;
  try {
    body = await parseBody(req);
  } catch (err) {
    return send(res, 400, { ok: false, message: '无法解析请求体：' + err.message });
  }

  const { customerName, customerContact, items, total } = body;
  if (!customerName || !customerContact) {
    return send(res, 400, { ok: false, message: '缺少客户名称或联系方式' });
  }
  if (!/^[\u4e00-\u9fa5·]{2,8}$/.test(customerName)) {
    return send(res, 400, { ok: false, message: '收货人必须为真实姓名（2-8个中文字符，可含中间点）' });
  }
  if (!/^\d{11}$/.test(customerContact)) {
    return send(res, 400, { ok: false, message: '联系电话必须为11位数字' });
  }
  if (!Array.isArray(items) || items.length === 0) {
    return send(res, 400, { ok: false, message: '缺少商品明细' });
  }

  const calcTotal = items.reduce((sum, it) => sum + Number(it.subtotal || 0), 0);
  if (Number(total || 0).toFixed(2) !== calcTotal.toFixed(2)) {
    return send(res, 400, { ok: false, message: '金额不一致，请重新提交' });
  }
  if (calcTotal < MIN_TOTAL) {
    return send(res, 400, { ok: false, message: `未达起送金额 ${MIN_TOTAL} 元` });
  }

  const record = {
    ...body,
    id: 'Q' + Date.now(),
    orderTime: body.orderTime || body.createdAt || new Date().toISOString(),
    receivedAt: new Date().toISOString(),
    total: calcTotal
  };

  // 生成 Excel 内容
  const excelData = buildExcel(record);
  const excelPath = path.join(EXCEL_DIR, excelData.filename);
  try {
    fs.mkdirSync(EXCEL_DIR, { recursive: true });
    fs.writeFileSync(excelPath, excelData.buffer);
    record.excelFilename = excelData.filename;
    record.excelPath = path.relative(__dirname, excelPath);
  } catch (err) {
    console.error('写入 Excel 失败', err);
  }

  fs.appendFile(LOG_FILE, JSON.stringify(record) + '\n', err => {
    if (err) {
      console.error('写入日志失败', err);
      return send(res, 500, { ok: false, message: '服务器写入失败' });
    }
    console.log(`收到确认单 ${record.id}，客户：${record.customerName}，金额：${record.total.toFixed(2)} 元`);
    send(res, 200, {
      ok: true,
      id: record.id,
      serverTime: record.receivedAt,
      record,
      excelFilename: excelData.filename,
      excelBase64: excelData.base64
    });
  });
}

const server = http.createServer((req, res) => {
  if (req.method === 'OPTIONS') {
    return send(res, 204, { ok: true });
  }

  // 简单路由解析
  const urlObj = new URL(req.url, 'http://localhost');
  const { pathname } = urlObj;

  // 下载 Excel：GET /api/huoguo-quote/:id/excel
  if (req.method === 'GET') {
    const match = pathname.match(/^\/api\/huoguo-quote\/([^/]+)\/excel$/);
    if (match) {
      const rawId = match[1];
      // 防止路径穿越
      const safeId = rawId.replace(/[^0-9A-Za-z_-]/g, '');
      const excelPath = path.join(EXCEL_DIR, `${safeId}.xlsx`);
      fs.stat(excelPath, (err, stats) => {
        if (err || !stats.isFile()) {
          return send(res, 404, { ok: false, message: '确认单不存在或尚未生成Excel' });
        }
        res.writeHead(200, {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'Content-Disposition': `attachment; filename="${safeId}.xlsx"`,
          'Access-Control-Allow-Origin': '*'
        });
        fs.createReadStream(excelPath).pipe(res);
      });
      return;
    }
  }

  if (req.method === 'POST' && pathname === '/api/huoguo-quote') {
    return handleQuote(req, res);
  }

  send(res, 404, { ok: false, message: 'Not Found' });
});

server.listen(PORT, () => {
  console.log(`Huoguo quote API listening on http://localhost:${PORT}`);
});

// 将确认单数据转成 Excel（xlsx）
function buildExcel(record) {
  const rows = [
    ['确认单编号', record.id],
    ['下单时间', record.orderTime],
    ['后台接收时间', record.receivedAt],
    ['客户名称', record.customerName],
    ['联系方式', record.customerContact],
    ['送货地址', record.customerAddress || '—'],
    ['总金额(元)', Number(record.total || 0)],
    [],
    ['序号', '商品名称', '规格', '单价(元)', '数量', '小计(元)']
  ];

  (record.items || []).forEach((item, idx) => {
    rows.push([
      idx + 1,
      item.name,
      item.spec,
      Number(item.price),
      Number(item.qty),
      Number(item.subtotal || 0)
    ]);
  });

  rows.push([]);
  rows.push(['确认金额', '', '', '', '', Number(record.total || 0)]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '确认单');

  const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  return {
    buffer,
    base64: buffer.toString('base64'),
    filename: `${record.id}.xlsx`
  };
}
