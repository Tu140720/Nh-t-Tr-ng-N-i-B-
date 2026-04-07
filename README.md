# Internal RAG Router

Ung dung nay hien thuc luong:

1. Nguoi dung dat cau hoi.
2. He thong tim trong `Internal Company Data Repository`.
3. Neu co du lieu phu hop, he thong di theo nhanh `RAG + AI xu ly` va tra loi.
4. Neu khong co du lieu, he thong fallback sang `Search Internet`.
5. AI tong hop ket qua va tra loi kem goi y.

## Cach chay

1. Tao file `.env` tu `.env.example`.
2. Neu muon bat AI va web search, them `GEMINI_API_KEY`.
3. Chay `npm start`.
   Neu PowerShell chan `npm.ps1`, dung `npm.cmd start` hoac `node server.mjs`.
4. Mo `http://127.0.0.1:<PORT>`.
5. Neu muon mo mini app danh cho Telegram, dung `http://127.0.0.1:<PORT>/telegram/`.
6. Neu muon gan mini app vao bot Telegram, them `TELEGRAM_BOT_TOKEN` va `TELEGRAM_MINI_APP_URL` vao `.env`, sau do chay `npm run telegram:set-menu`.

## Deploy Render

Repo da duoc them san:

- `Dockerfile`: dong goi app de chay tren Render.
- `render.yaml`: blueprint web service + persistent disk.
- `scripts/render-entrypoint.sh`: seed du lieu trong `data/` vao disk o lan deploy dau tien.

Luong deploy de dung duoc Telegram Mini App:

1. Day project len mot Git repo.
2. Tren Render, tao service tu `render.yaml`.
3. Dien cac env duoc danh dau `sync: false` trong dashboard.
4. Sau khi deploy xong, lay URL HTTPS cua service, vi du `https://ten-service.onrender.com`.
5. Cap nhat `.env` local:
   - `TELEGRAM_MINI_APP_URL=https://ten-service.onrender.com/telegram/`
6. Chay `npm run telegram:set-menu` tren may cua ban de gan nut mo mini app cho bot.

Neu ban muon domain rieng:

1. Vao service tren Render, them custom domain.
2. Cau hinh DNS theo huong dan Render.
3. Sau khi domain verify xong, doi `TELEGRAM_MINI_APP_URL` sang domain moi, vi du `https://miniapp.tenmiencuaban.com/telegram/`.

## Bien moi truong

- `PORT`: cong cua web app.
- `INTERNAL_DOCS_DIR`: thu muc chua tai lieu noi bo.
- `USERS_SHEET_URL`: link Google Sheet dung de dong bo danh sach user dang nhap.
- `USERS_SYNC_INTERVAL_MS`: chu ky tu dong dong bo user tu sheet, tinh bang mili giay.
- `ADMIN_OTP_TTL_MS`: thoi gian het han ma OTP email cho `admin`, tinh bang mili giay.
- `SMTP_HOST`, `SMTP_PORT`, `SMTP_SECURE`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM`: cau hinh SMTP de gui ma OTP email cho `admin`.
- `GEMINI_API_KEY`: khoa Gemini API de bat tong hop AI va web search.
- `GEMINI_BASE_URL`: mac dinh `https://generativelanguage.googleapis.com/v1beta`.
- `GEMINI_MODEL`: mac dinh `gemini-3-flash-preview`.
- `GEMINI_FALLBACK_MODELS`: danh sach model du phong cho nhanh noi bo, phan tach bang dau phay.
- `GEMINI_WEB_MODEL`: tuy chon, ghi de model cho nhanh web search.
- `GEMINI_WEB_FALLBACK_MODEL`: model fallback cho nhanh web khi model dau bi `429/5xx`, mac dinh `gemini-2.5-flash`.
- `GEMINI_WEB_FALLBACK_MODELS`: danh sach model du phong bo sung cho nhanh web, phan tach bang dau phay.
- `GEMINI_THINKING_LEVEL`: mac dinh `minimal`.
- `TELEGRAM_BOT_TOKEN`: bot token de cau hinh bot Telegram mo mini app.
- `TELEGRAM_MINI_APP_URL`: URL HTTPS tro den `/telegram/` tren server cua ban.
- `TELEGRAM_MENU_TEXT`: nhan tren nut menu bot, mac dinh `Mo mini app`.
- `TELEGRAM_CHAT_ID`: tuy chon, neu muon dat menu button cho mot chat private cu the thay vi mac dinh toan bot.

## Dac diem ky thuat

- Backend Node.js thuan, khong can framework ngoai.
- Truy hoi noi bo bang BM25 nhe tren cac chunk tai lieu trong `data/internal`.
- Khi co `GEMINI_API_KEY`, he thong goi `POST /v1beta/models/{model}:generateContent` cua Gemini.
- Nhanh Internet dung Google Search grounding cua Gemini.
- Nhanh noi bo va nhanh Internet deu ho tro retry qua nhieu model Gemini du phong khi gap `429/5xx`.
- UI cho phep them tai lieu noi bo truc tiep ngay tren trang.
- Co the dong bo user dang nhap tu Google Sheet; khi sheet doi, he thong se cap nhat lai user/quyen theo chu ky.
- Tai khoan co role `admin` se dang nhap qua 2 buoc: mat khau + ma OTP gui qua email trong sheet.

## Cau truc thu muc

- `server.mjs`: HTTP server va API.
- `lib/`: config, retrieval, fallback, Gemini client, internal repository.
- `public/`: giao dien.
- `public/telegram/`: mini app toi gian cho Telegram, chi gom 4 chuc nang chinh.
- `render.yaml`: cau hinh deploy Render.
- `scripts/render-entrypoint.sh`: khoi tao du lieu vao persistent disk o lan chay dau.
- `data/internal/`: du lieu noi bo mau.
- `test/`: test truy hoi.

## Cach test

1. Chay unit test bang `npm test`.
   Neu PowerShell chan `npm.ps1`, dung `node --test --test-isolation=none`.
2. Chay app bang `npm start` hoac `node server.mjs`.
3. Dang nhap tren web bang tai khoan dong bo tu Google Sheet.
   Mat khau mac dinh la `123456` cho toi khi nguoi dung tu doi mat khau.
4. Test nhanh tren giao dien:
   - Cau hoi noi bo: `Nhan vien chinh thuc duoc nghi phep nam bao nhieu ngay?`
   - Cau hoi web: `Gia vang the gioi hien nay dang tang hay giam?`
5. Co the test API truc tiep:
   - `POST /api/login`
   - `GET /api/health`
   - `POST /api/ask`
