"""Dars jadvali tuzuvchi: veb-ilova (Streamlit).

Ishga tushirish:  streamlit run app.py
Parol qo'yish (ixtiyoriy): .streamlit/secrets.toml ichida  PAROL = "sizning_parolingiz"
"""
import hmac
from pathlib import Path

import pandas as pd
import streamlit as st

from jadval_tuz import tuz_bytes

st.set_page_config(page_title="Dars jadvali tuzuvchi", page_icon="📅", layout="wide")

HERE = Path(__file__).parent


def parol_tekshir() -> bool:
    try:
        pw = st.secrets["PAROL"]
    except Exception:
        return True  # parol o'rnatilmagan: ochiq kirish
    if st.session_state.get("kirdi"):
        return True
    x = st.text_input("Parol", type="password")
    if x:
        if hmac.compare_digest(x, str(pw)):
            st.session_state["kirdi"] = True
            st.rerun()
        else:
            st.error("Parol noto'g'ri")
    return False


def jadval_df(data, events, res, kalit):
    days, pairs, groups, teachers, rooms, lessons = data
    D, P = len(days), len(pairs)
    cols = list(groups) if kalit == "guruh" else sorted({e["teacher"] for e in events})
    rows = []
    for d in range(D):
        for p in range(P):
            row = {"Kun": days[d][1], "Juftlik": f"{p + 1} ({pairs[p][1]})"}
            for c in cols:
                row[c] = "MUSTAQIL TA'LIM" if kalit == "guruh" and d in groups[c][1] else ""
            rows.append(row)
    for e in events:
        d, p, r = res[e["eid"]]
        if kalit == "guruh":
            txt = f"{e['subject']} ({e['type']}) · {e['teacher']} · {r}"
            targets = e["groups"]
        else:
            txt = f"{e['subject']} ({e['type']}) · {', '.join(e['groups'])} · {r}"
            targets = [e["teacher"]]
        for t in targets:
            rows[d * P + p][t] = txt
    return pd.DataFrame(rows)


if not parol_tekshir():
    st.stop()

st.title("📅 Dars jadvali tuzuvchi")
st.write(
    "1) Namuna Excel'ni yuklab oling va o'z ma'lumotlaringiz bilan to'ldiring. "
    "2) To'ldirilgan faylni yuklang. 3) «Jadval tuzish» tugmasini bosing."
)

namuna = HERE / "namuna_kirish.xlsx"
if namuna.exists():
    st.download_button("⬇️ Namuna kirish faylini yuklab olish", namuna.read_bytes(),
                       file_name="namuna_kirish.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with st.sidebar:
    st.header("Sozlamalar")
    vaqt = st.slider("Qidiruv vaqti (soniya)", 10, 300, 60, 10,
                     help="Uzoq vaqt = yaxshiroq jadval. Katta ma'lumotda oshiring.")
    st.caption("Muhimlik (0 = e'tiborsiz):")
    w_gap = st.slider("Guruhda «oyna»larni kamaytirish", 0, 20, 10)
    w_tgap = st.slider("O'qituvchida «oyna»larni kamaytirish", 0, 20, 3)
    w_over = st.slider("Kuniga 3 juftlikdan ortiq yuklamani kamaytirish", 0, 20, 3)
    w_late = st.slider("Darslarni erta juftliklarga yig'ish", 0, 5, 1)
    w_same = st.slider("Bir fan darslarini turli kunlarga yoyish", 0, 20, 5)

fayl = st.file_uploader("To'ldirilgan kirish fayli (.xlsx)", type=["xlsx"])

if fayl is not None and st.button("Jadval tuzish", type="primary"):
    try:
        with st.spinner("Jadval hisoblanmoqda..."):
            natija, holat, xatolar, tafsilot = tuz_bytes(
                fayl, vaqt, w_gap=w_gap, w_tgap=w_tgap, w_late=w_late,
                w_over=w_over, w_same=w_same)
        st.session_state["natija"] = (natija, holat, xatolar, tafsilot)
    except SystemExit as e:
        st.session_state.pop("natija", None)
        st.error(f"Jadval tuzilmadi: {e}")
    except KeyError as e:
        st.session_state.pop("natija", None)
        st.error(f"Kirish faylida «{e.args[0]}» varag'i topilmadi. Namuna fayl tuzilishidan foydalaning.")
    except Exception as e:  # noto'g'ri formatdagi fayl va h.k.
        st.session_state.pop("natija", None)
        st.error(f"Faylni o'qishda xato: {e}. Namuna fayl tuzilishiga mosligini tekshiring.")

if "natija" in st.session_state:
    natija, holat, xatolar, (data, events, res) = st.session_state["natija"]
    if xatolar:
        st.error("Tekshiruvda muammo topildi:\n\n" + "\n".join(f"- {x}" for x in xatolar))
    else:
        st.success(f"Jadval tayyor ({'eng yaxshi topildi' if holat == 'OPTIMAL' else 'yaxshi variant topildi; vaqtni oshirsangiz yaxshilanishi mumkin'}). "
                   "Tekshiruv: o'qituvchi, guruh va xona to'qnashuvi yo'q.")
    st.download_button("⬇️ Natijani Excel'da yuklab olish", natija, file_name="natija.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
    t1, t2 = st.tabs(["Guruhlar jadvali", "O'qituvchilar jadvali"])
    with t1:
        st.dataframe(jadval_df(data, events, res, "guruh"), use_container_width=True, hide_index=True)
    with t2:
        st.dataframe(jadval_df(data, events, res, "oqituvchi"), use_container_width=True, hide_index=True)
