# Dars jadvali tuzuvchi

Excel'dan ma'lumot o'qib, to'qnashuvsiz dars jadvalini avtomatik tuzadi (Google OR-Tools CP-SAT).

## Fayllar
- `app.py` — veb-ilova (Streamlit)
- `jadval_tuz.py` — jadval tuzuvchi yadro (terminaldan ham ishlaydi)
- `shablon_yarat.py` — namuna kirish faylini yaratadi
- `namuna_kirish.xlsx` — namuna kirish fayli
- `requirements.txt` — kerakli kutubxonalar

## Kompyuterda sinash
```
pip install -r requirements.txt
streamlit run app.py
```

## Internetga qo'yish (bepul, Streamlit Community Cloud)
1. GitHub'da yangi repository oching va shu papkadagi barcha fayllarni yuklang
   (`.streamlit/secrets.toml` ni YUKLAMANG).
2. https://share.streamlit.io ga GitHub akkauntingiz bilan kiring.
3. **Create app** → repository'ni tanlang → Main file: `app.py` → **Deploy**.
4. Bir necha daqiqadan keyin `https://....streamlit.app` havolasi tayyor bo'ladi. Shuni ulashasiz.
5. Parol qo'yish uchun: ilova sahifasida **Settings → Secrets** ga yozing:
   ```
   PAROL = "o'zingiz o'ylagan parol"
   ```
   Shundan keyin ilova ochilganda parol so'raydi.

## Eslatmalar
- Bepul serverda protsessor kuchsiz: katta ma'lumotda sidebar'dagi «Qidiruv vaqti»ni 120–300 soniyaga oshiring.
- Yuklangan fayllar serverda saqlanmaydi, faqat hisoblash vaqtida xotirada turadi.
- Ilova ochiq repository'dan ishlasa, repository'da haqiqiy o'qituvchi/guruh ma'lumotlarini saqlamang.
