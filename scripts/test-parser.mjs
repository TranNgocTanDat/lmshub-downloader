// Kiểm tra parseLink.ts bằng đoạn mã iframe thật của Tan (K5) + các trường hợp biên.
import fs from 'node:fs';
import { parseInput, repairDoubleDecoded } from '/tmp/parsetest/parseLink.js';

const IFRAME = `<iframe src="https://view.officeapps.live.com/op/embed.aspx?src=https%3A%2F%2Fmedia%2Elmshub%2Evn%3A443%2Flms%2Dprod%2Dcontent%2D593793041885%2Ftenants%2Fd9ca4268%2D19ac%2D40e7%2D85bf%2D790c4c191035%2Fcontents%2F328436e8%2D2d05%2D440e%2Db13f%2D078d0c2f074c%2FKh%25E1%25BB%2591i%25205%2D%2520Slide%2520%2D%2520Ki%25E1%25BB%2583m%2520ch%25E1%25BB%25A9ng%2520tin%2520%25C4%2591i%25E1%25BB%2593n%2520m%25E1%25BA%25A1ng%2520%2D%2520Mau%2520Am%2520Thanh%2520Voice%2520%2D%2520Sinh%2520dong%2520V3%2Epptx%3FX%2DAmz%2DAlgorithm%3DAWS4%2DHMAC%2DSHA256%26X%2DAmz%2DCredential%3DAKIAYUQGSYXOR3W3TCKV%252F20261120%252Fap%2Dsoutheast%2D1%252Fs3%252Faws4%5Frequest%26X%2DAmz%2DDate%3D20261120T090459Z%26X%2DAmz%2DExpires%3D3600%26X%2DAmz%2DSignedHeaders%3Dhost%26response%2Dcontent%2Dtype%3Dapplication%252Fvnd%2Eopenxmlformats%2Dofficedocument%2Epresentationml%2Epresentation%26X%2DAmz%2DSignature%3D951646ab317205260c18978a0832159eb8f84079344878ccda7973286a48f482&amp;wdAr=1.7777777777777777&amp;Embed=1" width="1200px" height="574px" frameborder="0" title="Trình xem PowerPoint"></iframe>`;

// Chuẩn đối chiếu: lấy thẳng từ mã iframe gốc bằng thư viện chuẩn (giải mã 1 lần).
// KHÔNG dùng /tmp/k5_url.txt nữa vì file đó do gõ tay lúc trước, thiếu chữ "i" trong "điồn".
const _srcRaw = IFRAME.match(/src="([^"]+)"/)[1].replace(/&amp;/g, '&');
const EXPECTED_URL = new URLSearchParams(_srcRaw.slice(_srcRaw.indexOf('?') + 1)).get('src');

let pass = 0;
let fail = 0;
function check(name, cond, extra = '') {
  if (cond) { pass++; console.log(`PASS ${name}`); }
  else { fail++; console.log(`FAIL ${name} ${extra}`); }
}

// 1. Mã iframe thật -> link trực tiếp phải KHỚP link đã tải được bằng curl.
const r1 = parseInput(IFRAME);
check('1. iframe thật -> ok', r1.ok === true);
if (r1.ok) {
  check('1a. directUrl khớp link curl', r1.result.directUrl === EXPECTED_URL, `\n  got: ${r1.result.directUrl}\n  exp: ${EXPECTED_URL}`);
  check('1b. filename tiếng Việt', r1.result.filename === 'Khối 5- Slide - Kiểm chứng tin điồn mạng - Mau Am Thanh Voice - Sinh dong V3.pptx', `got: ${r1.result.filename}`);
  check('1c. host', r1.result.host === 'media.lmshub.vn', `got: ${r1.result.host}`);
  check('1d. signed', r1.result.signed === true);
  check('1e. hết hạn 17:04:59 +07', r1.result.expiresAt?.toISOString() === '2026-11-20T10:04:59.000Z', `got: ${r1.result.expiresAt?.toISOString()}`);
}

// 2. Dán link embed trực tiếp (không có thẻ iframe) -> kết quả như trên.
const embedUrl = IFRAME.match(/src="([^"]+)"/)[1].replace(/&amp;/g, '&');
const r2 = parseInput(embedUrl);
check('2. link embed trực tiếp -> ok + khớp URL', r2.ok === true && r2.ok && r2.result.directUrl === EXPECTED_URL);

// 3. Dán link trực tiếp (đã giải mã 1 lần) -> giữ nguyên.
const r3 = parseInput(EXPECTED_URL);
check('3. link trực tiếp -> giữ nguyên', r3.ok === true && r3.ok && r3.result.directUrl === EXPECTED_URL);

// 4. Link bị giải mã 2 lần (chữ Việt thô + dấu / thô trong credential) -> tự sửa về đúng.
const doubleDecoded = decodeURIComponent(EXPECTED_URL);
const r4 = parseInput(doubleDecoded);
check('4. link giải mã 2 lần -> sửa về đúng', r4.ok === true && r4.ok && r4.result.directUrl === EXPECTED_URL, r4.ok ? `\n  got: ${r4.result.directUrl}` : `\n  err: ${r4.error}`);

// 5. repairDoubleDecoded không đổi gì với link đã đúng.
check('5. repair là no-op với link đúng', repairDoubleDecoded(EXPECTED_URL) === EXPECTED_URL);

// 6. Đầu vào rác / rỗng -> báo lỗi tiếng Việt.
const r6 = parseInput('hello world');
check('6a. chuỗi rác -> lỗi', r6.ok === false && !r6.ok && r6.error.length > 0);
const r7 = parseInput('   ');
check('6b. chuỗi rỗng -> lỗi', r7.ok === false);

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
