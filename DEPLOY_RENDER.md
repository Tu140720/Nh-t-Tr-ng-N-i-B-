# Deploy Render

## Muc tieu

Deploy app hien tai len Render, giu nguyen web app, mo them Telegram Mini App tai `/telegram/`, va seed du lieu hien co trong `data/` vao persistent disk o lan deploy dau.

## Canh bao

- Repo nen de `private`.
- Thu muc `data/` hien dang chua du lieu van hanh. Neu ban day len GitHub, toan bo du lieu nay se di cung repo.
- File `.env` da duoc ignore, khong day len repo.

## Da duoc chuan bi san

- `Dockerfile`
- `render.yaml`
- `scripts/render-entrypoint.sh`
- `scripts/telegram-set-menu.mjs`

## Cach lam

1. Tao mot Git repo private va push toan bo project nay len.
2. Tren Render, chon `New +` -> `Blueprint`.
3. Chon repo vua push.
4. Render se doc `render.yaml` va tao web service `ai-beta-001`.
5. Dien cac secret trong dashboard Render:
   - `USERS_SHEET_URL`
   - `SMTP_HOST`
   - `SMTP_USER`
   - `SMTP_PASS`
   - `SMTP_FROM`
   - `GEMINI_API_KEY`
6. Deploy service.
7. Sau khi xong, lay URL HTTPS cua service:
   - Vi du: `https://ai-beta-001.onrender.com`
8. Cap nhat `.env` local:
   - `TELEGRAM_MINI_APP_URL=https://ai-beta-001.onrender.com/telegram/`
9. Chay local:
   - `npm.cmd run telegram:set-menu`

## Domain rieng

1. Trong service Render, mo `Settings` -> `Custom Domains`.
2. Them domain, vi du:
   - `miniapp.tenmiencuaban.com`
3. Cau hinh DNS theo huong dan Render.
4. Cho Render cap SSL.
5. Cap nhat lai `.env` local:
   - `TELEGRAM_MINI_APP_URL=https://miniapp.tenmiencuaban.com/telegram/`
6. Chay lai:
   - `npm.cmd run telegram:set-menu`

## Kiem tra sau deploy

- `https://<domain>/api/health`
- `https://<domain>/telegram/`
- Dang nhap mini app bang tai khoan noi bo
- Tao don / theo doi don / hoan thanh giao hang
