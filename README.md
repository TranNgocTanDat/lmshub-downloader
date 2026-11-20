# Lấy link tải slide lmshub

Trang React + Vite + TypeScript: dán đoạn mã iframe trong hộp thoại
"Nhúng Tài liệu" của lmshub vào là lấy được link tải trực tiếp của file slide.

## Chạy

```bash
npm install
npm run dev      # chạy thử ở http://localhost:5173
npm run build    # build ra thư mục dist/
```

## Cách nó hoạt động

Link file thật nằm trong tham số `src=` của link xem trước
(`view.officeapps.live.com/op/embed.aspx?src=...`) và đang bị mã hoá URL.
Trang lấy giá trị `src`, giải mã **đúng 1 lần** là ra link tải trực tiếp
(`media.lmshub.vn/...?X-Amz-Signature=...`).

Quy tắc quan trọng nhất: **chỉ giải mã 1 lần**. Giải mã 2 lần (hoặc giải mã
tay gõ nhầm 1 ký tự) làm hỏng chữ ký S3 -> lỗi `SignatureDoesNotMatch`.
Vì vậy code dùng `URLSearchParams.get('src')` (nó đã giải mã 1 lần) và không
giải mã thêm nữa. Nếu người dùng lỡ dán link đã bị giải mã 2 lần, hàm
`repairDoubleDecoded` trong `src/lib/parseLink.ts` sẽ tự sửa lại
(chữ Việt thô -> %XX, dấu `/` thô trong credential -> %2F, khoảng trắng
thô trong đường dẫn -> %20).

## Kiểm tra

```bash
node scripts/test-parser.mjs   # 12 test, dùng đúng mã iframe thật (K5)
```

Test dùng mã iframe thật của bài "Khối 5 - Kiểm chứng tin điồn mạng" và kiểm
tra link ra khớp đúng link đã tải được bằng curl (40 MB, 26 slides), kèm hạn
dùng tính từ `X-Amz-Date` + `X-Amz-Expires`.

## Lưu ý

- Link tải chỉ sống khoảng 1 giờ kể từ lúc mở hộp thoại "Nhúng Tài liệu".
  Quá hạn thì mở lại xem trước, lấy mã mới.
- Nút "Tải xuống file" mở thẳng link tải trực tiếp. Đừng dán link đó vào
  thanh địa chỉ để xem trước, trình xem Office sẽ báo không tìm thấy file.
