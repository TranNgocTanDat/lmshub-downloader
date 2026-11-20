import { useEffect, useMemo, useState } from 'react'
import { parseInput, type ParsedLink } from './lib/parseLink'

const SAMPLE_IFRAME = `<iframe src="https://view.officeapps.live.com/op/embed.aspx?src=https%3A%2F%2Fmedia%2Elmshub%2Evn%3A443%2Flms%2Dprod%2Dcontent%2D593793041885%2Ftenants%2Fd9ca4268%2D19ac%2D40e7%2D85bf%2D790c4c191035%2Fcontents%2F328436e8%2D2d05%2D440e%2Db13f%2D078d0c2f074c%2FKh%25E1%25BB%2591i%25205%2D%2520Slide%2520%2D%2520Ki%25E1%25BB%2583m%2520ch%25E1%25BB%25A9ng%2520tin%2520%25C4%2591i%25E1%25BB%2593n%2520m%25E1%25BA%25A1ng%2520%2D%2520Mau%2520Am%2520Thanh%2520Voice%2520%2D%2520Sinh%2520dong%2520V3%2Epptx%3FX%2DAmz%2DAlgorithm%3DAWS4%2DHMAC%2DSHA256%26X%2DAmz%2DCredential%3DAKIAYUQGSYXOR3W3TCKV%252F20261120%252Fap%2Dsoutheast%2D1%252Fs3%252Faws4%5Frequest%26X%2DAmz%2DDate%3D20261120T090459Z%26X%2DAmz%2DExpires%3D3600%26X%2DAmz%2DSignedHeaders%3Dhost%26response%2Dcontent%2Dtype%3Dapplication%252Fvnd%2Eopenxmlformats%2Dofficedocument%2Epresentationml%2Epresentation%26X%2DAmz%2DSignature%3D951646ab317205260c18978a0832159eb8f84079344878ccda7973286a48f482&amp;wdAr=1.7777777777777777&amp;Embed=1" width="1200px" height="574px" frameborder="0"></iframe>`

function formatRemaining(expiresAt: Date | null, now: number): string {
  if (!expiresAt) return 'Không rõ hạn dùng'
  const ms = expiresAt.getTime() - now
  if (ms <= 0) return 'Đã hết hạn'
  const mins = Math.floor(ms / 60000)
  if (mins < 1) return 'Còn dưới 1 phút'
  if (mins < 60) return `Còn khoảng ${mins} phút`
  const h = Math.floor(mins / 60)
  const m = mins % 60
  return `Còn khoảng ${h} giờ ${m} phút`
}

export default function App() {
  const [input, setInput] = useState('')
  const [result, setResult] = useState<ParsedLink | null>(null)
  const [error, setError] = useState('')
  const [copied, setCopied] = useState(false)
  const [now, setNow] = useState(Date.now())

  // Đồng hồ đếm ngược hạn link, cập nhật mỗi 30 giây.
  useEffect(() => {
    const id = window.setInterval(() => setNow(Date.now()), 30000)
    return () => window.clearInterval(id)
  }, [])

  const expired = useMemo(
    () => (result?.expiresAt ? result.expiresAt.getTime() <= now : false),
    [result, now],
  )

  function handleParse(text: string) {
    const parsed = parseInput(text)
    if (parsed.ok) {
      setResult(parsed.result)
      setError('')
    } else {
      setResult(null)
      setError(parsed.error)
    }
    setCopied(false)
  }

  function handleInputChange(text: string) {
    setInput(text)
    // Tự phân tích khi dán xong đoạn mã/link đầy đủ.
    if (text.includes('</iframe>') || text.includes('X-Amz-Signature')) {
      handleParse(text)
    } else {
      setResult(null)
      setError('')
    }
  }

  async function handleCopy() {
    if (!result) return
    try {
      await navigator.clipboard.writeText(result.directUrl)
      setCopied(true)
      window.setTimeout(() => setCopied(false), 2000)
    } catch {
      setError('Trình duyệt không cho sao chép tự động, bạn bôi đen link rồi bấm Ctrl+C nhé.')
    }
  }

  return (
    <main className="page">
      <header>
        <h1>Lấy link tải slide lmshub</h1>
        <p>
          Dán đoạn mã iframe trong hộp thoại <strong>“Nhúng Tài liệu”</strong> vào, trang sẽ
          lấy link tải trực tiếp cho bạn. Link chỉ có hiệu lực khoảng 1 giờ nên lấy xong thì
          tải ngay.
        </p>
      </header>

      <section className="card">
        <label htmlFor="paste">Dán mã iframe hoặc link vào đây</label>
        <textarea
          id="paste"
          value={input}
          onChange={(e) => handleInputChange(e.target.value)}
          placeholder='<iframe src="https://view.officeapps.live.com/op/embed.aspx?src=..." ...></iframe>'
          rows={5}
          spellCheck={false}
        />
        <div className="actions">
          <button type="button" className="primary" onClick={() => handleParse(input)}>
            Lấy link tải
          </button>
          <button
            type="button"
            onClick={() => {
              setInput(SAMPLE_IFRAME)
              handleParse(SAMPLE_IFRAME)
            }}
          >
            Dán mã mẫu thử
          </button>
          <button
            type="button"
            onClick={() => {
              setInput('')
              setResult(null)
              setError('')
            }}
          >
            Xoá
          </button>
        </div>
        {error && <p className="error">{error}</p>}
      </section>

      {result && (
        <section className="card result">
          <h2>{result.filename}</h2>
          <dl>
            <div>
              <dt>Nguồn</dt>
              <dd>{result.host || '—'}</dd>
            </div>
            <div>
              <dt>Hạn dùng</dt>
              <dd className={expired ? 'expired' : 'ok'}>
                {formatRemaining(result.expiresAt, now)}
                {result.expiresAt &&
                  ` (hết hạn lúc ${result.expiresAt.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' })})`}
              </dd>
            </div>
            <div>
              <dt>Chữ ký</dt>
              <dd>{result.signed ? 'Có (link đầy đủ)' : 'Không thấy chữ ký — có thể tải lỗi'}</dd>
            </div>
          </dl>

          {expired && (
            <p className="error">
              Link này đã hết hạn. Bạn mở lại xem trước slide, bấm “Thông tin Nhúng” lấy mã mới
              rồi dán vào đây.
            </p>
          )}

          <label htmlFor="direct">Link tải trực tiếp</label>
          <input id="direct" type="text" readOnly value={result.directUrl} onFocus={(e) => e.target.select()} />
          <div className="actions">
            <a
              className="primary download"
              href={result.directUrl}
              target="_blank"
              rel="noopener noreferrer"
            >
              Tải xuống file
            </a>
            <button type="button" onClick={handleCopy}>
              {copied ? 'Đã sao chép' : 'Sao chép link'}
            </button>
          </div>
          <p className="hint">
            Bấm “Tải xuống file” là trình duyệt tải thẳng file về. Đừng dán link này vào thanh
            địa chỉ để xem trước, trình xem Office sẽ báo không tìm thấy file.
          </p>
        </section>
      )}

      <section className="card guide">
        <h2>Lấy mã iframe ở đâu?</h2>
        <ol>
          <li>Mở xem trước slide trong trang Quản lý Nội dung lmshub.</li>
          <li>
            Ở thanh xám cuối màn hình xem trước, bấm nút <strong>“…”</strong> bên trái nút toàn
            màn hình, rồi chọn <strong>“Thông tin Nhúng”</strong>. Nếu bấm “…” không mở menu thì
            bấm nút ▶ sang slide sau, rồi bấm Tab một lần, Enter một lần.
          </li>
          <li>
            Hộp thoại “Nhúng Tài liệu” hiện ra với đoạn mã đã được bôi đen sẵn — copy toàn bộ
            đoạn mã đó rồi dán vào ô trên.
          </li>
        </ol>
        <p className="hint">
          Lưu ý: link tải chỉ sống khoảng 1 giờ kể từ lúc mở hộp thoại “Nhúng Tài liệu”. Quá
          hạn thì lấy mã mới. Trang này chỉ giải mã đúng 1 lần nên không làm hỏng chữ ký.
        </p>
      </section>
    </main>
  )
}
