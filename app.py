import streamlit as st
import pandas as pd
import gspread
from google.oauth2 import service_account
from io import BytesIO

st.set_page_config(page_title="X5 Торги — Автозаполнение", layout="centered")
st.title("📦 Автозаполнение ставок для торгов X5")

# === 1. Загрузка справочника из Google Sheets ===
@st.cache_data(ttl=3600)
def load_prices_from_gsheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scope
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key("17q72JGZmLfGcKjMcWhj8RBdnq06dsUFkmeZi6kuaXpA").worksheet("Цены")
    
    all_values = sheet.get_all_values()
    data_rows = all_values[3:]
    
    selected = []
    for row in data_rows:
        if len(row) < 16:
            row = row + [''] * (16 - len(row))
        selected.append([
            row[0],   # PLU
            row[2],   # Страна
            row[5],   # F
            row[6],   # G
            row[7],   # H
            row[8],   # I
            row[9],   # J
            row[10],  # K
            row[11],  # L
            row[12],  # M
            row[13],  # N
            row[14],  # O
            row[15]   # P
        ])
    
    col_names = ["PLU", "Страна", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
    df = pd.DataFrame(selected, columns=col_names)
    df = df[df["PLU"].notna() & (df["PLU"] != "")]
    df["PLU"] = df["PLU"].astype(str).str.strip()
    
    for col in ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    
    return df

# === 2. Справочник РЦ → колонка ===
RC_TO_COLUMN = {
    "Софьино": "F",
    "Воронеж": "G",
    "Рамонь": "G",
    "Курск": "G",
    "Тамбов": "G",
    "Саратов": "H",
    "Самара": "H",
    "Кузнецк": "H",
    "Волгоград": "H",
    "СЛК": "H",
    "Санкт-Петербург": "H",
    "Нижний Новгород": "H",
    "Адыгея": "I",
    "Невинномысск": "I",
    "Краснодар": "I",
    "Ростов": "I",
    "УФА": "J",
    "Оренбург": "J",
    "Пермь": "J",
    "Падиково": "K",
    "Литвиново": "K",
    "Валищево": "K",
    "Купавна": "K",
    "Ярославль": "L",
    "Дзержинск": "L",
    "Казань": "M",
    "Ростов-на-Дону": "N",
    "Екатеринбург": "O",
    "Челябинск": "O",
    "Кемерово": "P",
    "Новосибирск": "P",
}

def get_price_column_for_rc(rc_name):
    for key, col in RC_TO_COLUMN.items():
        if key.lower() in rc_name.lower():
            return col
    return None

# === 3. Загрузка файла ===
uploaded_file = st.file_uploader(
    "Загрузите файл с торгами (Excel)",
    type=["xlsx", "xls"],
    help="Скачайте файл с портала Х5 и загрузите его сюда"
)

# === 4. Параметры ===
col1, col2 = st.columns(2)
with col1:
    comment = st.text_input("📝 Комментарий (общий для всех позиций)", placeholder="Оставьте пустым, чтобы не менять")
with col2:
    discount = st.number_input("💰 Скидка (руб.)", min_value=0.0, step=0.5, value=0.0, help="0 — без скидки")

# === 5. Обработка ===
if uploaded_file is not None and st.button("🚀 Обработать файл"):
    with st.spinner("Идёт обработка..."):
        try:
            prices_df = load_prices_from_gsheets()
        except Exception as e:
            st.error(f"Ошибка при загрузке Google-таблицы: {e}")
            st.stop()

        try:
            df = pd.read_excel(uploaded_file, sheet_name=0)
        except Exception as e:
            st.error(f"Ошибка при чтении файла: {e}")
            st.stop()

        # Приводим текстовые колонки к строке, числовые — к числу
        df["Страна"] = df["Страна"].astype(str)
        df["Мой комментарий"] = df["Мой комментарий"].astype(str)
        df["Мое предложение"] = pd.to_numeric(df["Мое предложение"], errors="coerce")
        df["Мой гарантированный объем"] = pd.to_numeric(df["Мой гарантированный объем"], errors="coerce")

        # Временные колонки
        df["Цена_из_справочника"] = None
        df["Страна_из_справочника"] = None
        unknown_rcs = set()

        for idx, row in df.iterrows():
            plu = str(row["Код PLU"]).strip()
            rc = row["РЦ доставки"]

            price_row = prices_df[prices_df["PLU"] == plu]
            if price_row.empty:
                continue

            col_letter = get_price_column_for_rc(rc)
            if col_letter is None:
                unknown_rcs.add(rc)
                continue

            price_value = price_row.iloc[0][col_letter]
            if pd.notna(price_value):
                df.at[idx, "Цена_из_справочника"] = float(price_value)
                country = price_row.iloc[0]["Страна"]
                if pd.isna(country) or country is None:
                    country = ""
                df.at[idx, "Страна_из_справочника"] = str(country)

        if unknown_rcs:
            st.warning(f"⚠️ Для следующих РЦ не найдено соответствие: {', '.join(unknown_rcs)}")

        updated_count = 0
        for idx, row in df.iterrows():
            if pd.notna(row.get("Цена_из_справочника")):
                new_price = row["Цена_из_справочника"] - discount
                df.at[idx, "Мое предложение"] = round(new_price, 2)
                df.at[idx, "Страна"] = row["Страна_из_справочника"]
                if comment.strip():
                    df.at[idx, "Мой комментарий"] = comment.strip()
                df.at[idx, "Мой гарантированный объем"] = row["Количество"]
                updated_count += 1

        df = df.drop(columns=["Цена_из_справочника", "Страна_из_справочника"], errors="ignore")

        # Удаляем строки, где нет цены или страны
        before = len(df)
        df = df[df["Мое предложение"].notna() & df["Страна"].notna()]
        after = len(df)
        deleted = before - after

        if df.empty:
            st.warning("⚠️ После удаления пустых позиций не осталось ни одной строки.")
            st.stop()

        # Заменяем NaN на пустые строки во всех колонках (кроме F и M, которые мы очистим отдельно)
        df = df.fillna("")

        # Очищаем столбцы F и M (делаем полностью пустыми)
        # Внимание: названия столбцов должны точно совпадать с теми, что есть в файле
        if "Предложение X5 с доставкой" in df.columns:
            df["Предложение X5 с доставкой"] = None
        else:
            # Если название другое (например, после переименования), можно использовать индекс
            # Но будем считать, что название стандартное
            st.warning("⚠️ Столбец 'Предложение X5 с доставкой' не найден, пропускаем очистку.")
        
        if "Мое предложение (самовывоз)" in df.columns:
            df["Мое предложение (самовывоз)"] = None
        else:
            st.warning("⚠️ Столбец 'Мое предложение (самовывоз)' не найден, пропускаем очистку.")

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Сбор предложений")
        output.seek(0)

        st.success(f"✅ Обновлено {updated_count} позиций. Удалено пустых: {deleted}.")
        if deleted > 0:
            st.info(f"ℹ️ Удалены строки, которые не были заполнены ни автоматически, ни вручную.")

        st.download_button(
            label="📥 Скачать файл",
            data=output,
            file_name=uploaded_file.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        with st.expander("👁️ Превью результата (первые 5 строк)"):
            st.dataframe(df.head())