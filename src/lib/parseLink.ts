/**
 * Phân tích mã iframe / link từ lmshub ("Thông tin Nhúng") thành link tải trực tiếp.
 *
 * Quy tắc quan trọng nhất: link file thật nằm trong tham số `src=` của link xem
 * trước (view.officeapps.live.com) và đang bị mã hoá URL. Chỉ giải mã ĐÚNG 1 LẦN.
 * Giải mã 2 lần (hoặc giải mã tay gõ nhầm 1 ký tự) làm hỏng chữ ký S3
 * -> lỗi SignatureDoesNotMatch.
 */

export interface ParsedLink {
  /** Link tải trực tiếp, dùng nguyên làm href (không mã hoá lại). */
  directUrl: string;
  /** Tên file để hiển thị (đã giải mã tiếng Việt). */
  filename: string;
  host: string;
  /** Thời điểm link hết hạn (từ X-Amz-Date + X-Amz-Expires), null nếu không có. */
  expiresAt: Date | null;
  /** Có chữ ký S3 kèm theo không. */
  signed: boolean;
}

export type ParseResult =
  | { ok: true; result: ParsedLink }
  | { ok: false; error: string };

function decodeOnce(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

/** Mã hoá chỉ những ký tự non-ASCII còn thô, giữ nguyên các chuỗi %XX đã có. */
function encodeRawNonAscii(s: string): string {
  let out = '';
  for (const ch of s) {
    out += (ch.codePointAt(0) ?? 0) > 127 ? encodeURIComponent(ch) : ch;
  }
  return out;
}

/**
 * Sửa link presigned bị giải mã 2 lần (tên file còn chữ Việt thô, credential còn
 * dấu `/` thô thay vì %2F). Với link đã đúng (giải mã 1 lần) thì hàm này không
 * đổi gì cả.
 */
export function repairDoubleDecoded(url: string): string {
  const qIndex = url.indexOf('?');
  const path = qIndex >= 0 ? url.slice(0, qIndex) : url;
  const query = qIndex >= 0 ? url.slice(qIndex + 1) : '';
  // Khoảng trắng thô trong đường dẫn chỉ có thể do giải mã quá tay (%20 -> ' '),
  // mã hoá lại thành %20 để khớp đúng link đã ký.
  const fixedPath = encodeRawNonAscii(path).split(' ').join('%20');
  if (!query) return fixedPath;

  const fixedQuery = query
    .split('&')
    .map((pair) => {
      const eq = pair.indexOf('=');
      if (eq < 0) return encodeRawNonAscii(pair);
      const key = pair.slice(0, eq);
      let value = pair.slice(eq + 1);
      // Hai tham số này không được chứa dấu `/` thô trong chữ ký.
      if ((key === 'X-Amz-Credential' || key === 'response-content-type') && value.includes('/')) {
        value = value.split('/').join('%2F');
      }
      value = encodeRawNonAscii(value);
      return `${key}=${value}`;
    })
    .join('&');
  return `${fixedPath}?${fixedQuery}`;
}

function parseAmzDate(value: string | null): Date | null {
  if (!value) return null;
  const m = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$/.exec(value);
  if (!m) return null;
  return new Date(
    Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]),
  );
}

function buildResult(directUrl: string): ParsedLink {
  const clean = repairDoubleDecoded(directUrl);
  let filename = 'file';
  let host = '';
  let expiresAt: Date | null = null;
  let signed = false;
  try {
    const u = new URL(clean);
    host = u.hostname;
    const last = u.pathname.split('/').filter(Boolean).pop() ?? '';
    if (last) filename = decodeOnce(last);
    const params = new URLSearchParams(u.search);
    signed = params.has('X-Amz-Signature');
    const date = parseAmzDate(params.get('X-Amz-Date'));
    const expires = Number(params.get('X-Amz-Expires') ?? '');
    if (date && Number.isFinite(expires) && expires > 0) {
      expiresAt = new Date(date.getTime() + expires * 1000);
    }
  } catch {
    // Giữ giá trị mặc định, link vẫn trả về để người dùng tự thử.
  }
  return { directUrl: clean, filename, host, expiresAt, signed };
}

/**
 * Lấy giá trị tham số `src` trong link xem trước của Office.
 * URLSearchParams.get đã giải mã 1 lần rồi (đúng 1 lần cần thiết) nên KHÔNG
 * được giải mã thêm nữa, nếu không `%20` thành khoảng trắng và hỏng chữ ký.
 */
function extractSrcParam(embedUrl: string): string | null {
  const qIndex = embedUrl.indexOf('?');
  if (qIndex < 0) return null;
  const params = new URLSearchParams(embedUrl.slice(qIndex + 1));
  const src = params.get('src');
  return src || null;
}

export function parseInput(rawInput: string): ParseResult {
  const input = rawInput.trim();
  if (!input) {
    return { ok: false, error: 'Bạn dán đoạn mã iframe hoặc link vào ô trên trước đã.' };
  }

  // Mã iframe trong hộp thoại "Nhúng Tài liệu" dùng &amp; thay cho &.
  const text = input.replace(/&amp;/g, '&');

  // 1. Dán nguyên đoạn mã <iframe ...> -> lấy thuộc tính src.
  const iframeMatch = /<iframe[^>]*\ssrc\s*=\s*["']([^"']+)["']/i.exec(text);
  if (iframeMatch) {
    const srcValue = iframeMatch[1];
    if (srcValue.includes('view.officeapps.live.com')) {
      const direct = extractSrcParam(srcValue);
      if (!direct) {
        return { ok: false, error: 'Không tìm thấy tham số src= trong link xem trước. Bạn copy lại toàn bộ đoạn mã trong hộp thoại "Nhúng Tài liệu" rồi dán vào nhé.' };
      }
      return { ok: true, result: buildResult(direct) };
    }
    // src trỏ thẳng tới file (hiếm) thì dùng luôn.
    return { ok: true, result: buildResult(srcValue) };
  }

  // 2. Dán link xem trước của Office (không có thẻ iframe).
  if (text.includes('view.officeapps.live.com')) {
    const direct = extractSrcParam(text);
    if (!direct) {
      return { ok: false, error: 'Link xem trước này thiếu tham số src=. Bạn copy lại đoạn mã trong hộp thoại "Nhúng Tài liệu" rồi dán vào nhé.' };
    }
    return { ok: true, result: buildResult(direct) };
  }

  // 3. Dán link file còn đang mã hoá (bắt đầu bằng https%3A%2F%2F...) -> giải mã 1 lần.
  if (/^https?%3A/i.test(text)) {
    return { ok: true, result: buildResult(decodeOnce(text)) };
  }

  // 4. Dán link file trực tiếp (đã giải mã 1 lần) -> dùng nguyên, chỉ sửa nếu bị giải mã 2 lần.
  if (/^https?:\/\//i.test(text)) {
    return { ok: true, result: buildResult(text) };
  }

  return {
    ok: false,
    error: 'Không nhận ra đây là mã iframe hay link. Bạn dán nguyên đoạn mã trong hộp thoại "Nhúng Tài liệu" (bắt đầu bằng <iframe) vào nhé.',
  };
}
