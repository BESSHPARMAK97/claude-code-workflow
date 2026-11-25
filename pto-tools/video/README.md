# Сборка видео-инструкции

Нужны: Node + Playwright, Python 3 + Pillow + openpyxl, ffmpeg, RHVoice с русскими голосами
(`apt-get install rhvoice rhvoice-russian`).

1. `node video/shots.js video/work` — скриншоты генератора (берёт `dist/Генератор_АОСР.html`) и комплект Word в `video/work/assets/`.
   Перед этим положите в `video/work/assets/` файлы `реестр_работ.xlsx` и `реестр_с_ошибками.xlsx`.
2. Отрисуйте комплект в PNG (`soffice --convert-to pdf`, затем `pdftoppm -r 130 -png … kit`).
3. `python3 video/sheet_mock.py && node video/sheetshots.js video/work` — кадры листов реестра.
4. `python3 video/build.py` — кадры, озвучка, сборка `dist/Видео_инструкция_АОСР.mp4` с субтитрами.

Сценарий и текст озвучки — `video/scenes.py`. Голос задаётся в `build.py` (`VOICE`).
