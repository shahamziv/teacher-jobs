import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="עיבוד משרות מורים", layout="wide")
st.title("📄 עיבוד אוטומטי של קובץ משרות")

uploaded_file = st.file_uploader("העלה קובץ Excel של משרות:", type=["xlsx"])

def find_column(cols, keywords):
    for kw in keywords:
        for col in cols:
            if pd.isna(col):
                continue
            if kw in str(col):
                return col
    return None

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    teacher_sheets = [s for s in xls.sheet_names if s not in ["מחנכים", "בעלי תפקידים", "מאזן שעות"]]
    processed_rows = []

    known_subjects = {'ברגיל מאור': 'אנגלית', 'כספר טל': 'אנגלית', 'פורת אורלי': 'אנגלית', 'ארונוב אולגה': 'אנגלית', 'חנצ׳ינסקי נעמי': 'אנגלית', 'רונן מעיין': 'אנגלית', 'לנדסמן פיפרנו נתי': 'אנגלית', 'לב נטע': 'אנגלית', 'דגן לילך': 'אנגלית', 'אייזיקסון מירב': 'אנגלית', 'ברוק סיגל': 'אנגלית', 'באואר מעיין': 'אנגלית', 'נקר תמי': 'אנגלית', 'אנה קוטוזובה': 'אנגלית', 'שוחט איילת': 'אנגלית', 'עבד אלחלים דיאנא': 'אנגלית', 'לוז רחל': 'מתמטיקה', 'אבינרי אנג׳לינה': 'מתמטיקה', 'כהן ענבל': 'מתמטיקה', 'מור יעל': 'מתמטיקה', 'מנשה חיים': 'מתמטיקה', 'דיין רותם': 'מתמטיקה'}
'אבינרי אנגֲלינה': 'מתמטיקה',
'שוחט איילת': 'אנגלית',
'כהן ענבל': 'מתמטיקה'
    }

    for sheet in teacher_sheets:
        try:
            raw = xls.parse(sheet, header=None)
            teacher_name = str(raw.iloc[0, 1]).strip() if pd.notna(raw.iloc[0, 1]) else sheet

            df = xls.parse(sheet, header=3)
            df.columns = df.columns.astype(str)

            col_kita = find_column(df.columns, ['כיתה'])
            col_miktzoa = find_column(df.columns, ['מקצוע', 'תחום'])
            if not col_kita or not col_miktzoa:
                continue

            base = df[[col_kita, col_miktzoa]].copy()
            base.columns = ['כיתה', 'מקצוע']

            col_opec = find_column(df.columns, ['אופק'])
            col_oz = find_column(df.columns, ['עוז'])
            hours = pd.to_numeric(df.get(col_opec), errors='coerce').fillna(0) +                     pd.to_numeric(df.get(col_oz), errors='coerce').fillna(0)

            cols_role = [c for c in df.columns if any(x in c for x in ['ריכוז', 'תפקיד'])]
            role_hours = pd.DataFrame()
            for c in cols_role:
                role_hours[c] = pd.to_numeric(df[c], errors='coerce')
            hours_from_roles = role_hours.max(axis=1).fillna(0)
            hours[hours_from_roles > 0] = hours_from_roles[hours_from_roles > 0]

            gmul_cols = [c for c in df.columns if any(x in c for x in ['חינוך', 'בגרות', 'ריכוז', 'תפקיד'])]
            gmul_total = df[gmul_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)

            base['שם מורה'] = teacher_name
            base['מספר שעות'] = hours
            base['שעות גמול'] = gmul_total
            base['גליון'] = sheet

            base = base[base['מקצוע'].notna()]
            base = base[~base['מקצוע'].astype(str).str.contains('גמול.*%|סה\"כ גמולים|^0$', regex=True)]

            ranks = ['3 יחל', '4 יחל', '5 יחל', '3/4 יחל', '4/5 יחל']
            def interpret_subject_and_level(row):
                text = str(row['מקצוע']).strip()
                if text in ranks:
                    inferred = known_subjects.get(row['שם מורה'], None)
                    return pd.Series([inferred if inferred else text, text])
                return pd.Series([text, None])

            base[['מקצוע', 'רמה']] = base.apply(interpret_subject_and_level, axis=1)
            processed_rows.append(base[['שם מורה', 'מקצוע', 'רמה', 'כיתה', 'מספר שעות', 'שעות גמול', 'גליון']])

        except Exception:
            continue

    if processed_rows:
        final_df = pd.concat(processed_rows, ignore_index=True)
        st.success(f"✔️ נמצאו {len(final_df)} שורות תקינות")
        st.dataframe(final_df, use_container_width=True)

        buffer = io.BytesIO()
        final_df.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)

        st.download_button(
            label="🔹 הורד את הקובץ המעובד",
            data=buffer,
            file_name="teacher_jobs_cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ לא נמצאו נתונים תקינים לעיבוד.")