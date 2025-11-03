import streamlit as st
import pandas as pd
import re
from math import ceil
import io

st.set_page_config(page_title="Wine2WeShip", page_icon="🍷", layout="centered")

st.title("🍷 Wine2WeShip – CSV → XLSX per spedizioni USA")

st.write(
    "Carica il CSV, l’app filtrerà solo le righe con **United States of America**, "
    "dividerà in colli da 12, calcolerà il peso e userà le colonne del template "
    "**weshipbase.xlsx** presente nel repo."
)

TEMPLATE_PATH = "weshipbase.xlsx"   # <--- nuovo nome, in xlsx

# --- helper per gli stati USA ---
STATE_TO_ABBR = {
    'Alabama': 'AL','Alaska': 'AK','Arizona': 'AZ','Arkansas': 'AR',
    'California': 'CA','Colorado': 'CO','Connecticut': 'CT','Delaware': 'DE',
    'District of Columbia': 'DC','Washington DC': 'DC','Washington, DC': 'DC','DC': 'DC',
    'Florida': 'FL','Georgia': 'GA','Hawaii': 'HI','Idaho': 'ID','Illinois': 'IL',
    'Indiana': 'IN','Iowa': 'IA','Kansas': 'KS','Kentucky': 'KY','Louisiana': 'LA',
    'Maine': 'ME','Maryland': 'MD','Massachusetts': 'MA','Michigan': 'MI','Minnesota': 'MN',
    'Mississippi': 'MS','Missouri': 'MO','Montana': 'MT','Nebraska': 'NE','Nevada': 'NV',
    'New Hampshire': 'NH','New Jersey': 'NJ','New Mexico': 'NM','New York': 'NY',
    'North Carolina': 'NC','North Dakota': 'ND','Ohio': 'OH','Oklahoma': 'OK','Oregon': 'OR',
    'Pennsylvania': 'PA','Rhode Island': 'RI','South Carolina': 'SC','South Dakota': 'SD',
    'Tennessee': 'TN','Texas': 'TX','Utah': 'UT','Vermont': 'VT','Virginia': 'VA',
    'Washington': 'WA','West Virginia': 'WV','Wisconsin': 'WI','Wyoming': 'WY',
    'Puerto Rico': 'PR','Guam': 'GU','American Samoa': 'AS','U.S. Virgin Islands': 'VI','Northern Mariana Islands': 'MP'
}
ABBR_SET = set(STATE_TO_ABBR.values())


def normalize_state(value: str) -> str:
    if not isinstance(value, str):
        return ""
    s = value.strip()
    if not s:
        return ""
    if len(s) == 2 and s.upper() in ABBR_SET:
        return s.upper()
    s_clean = re.sub(r'[^A-Za-z ]+', '', s).strip()
    variants = {'NYC': 'NY', 'DISTRITO OF COLUMBIA': 'DC', 'WASHINGTON DC': 'DC', 'WASHINGTON': 'WA'}
    if s_clean.upper() in variants:
        return variants[s_clean.upper()]
    abbr = STATE_TO_ABBR.get(s_clean.title())
    if abbr:
        return abbr
    tokens = [t.strip() for t in re.split(r'[,/|-]+', s)]
    for t in reversed(tokens):
        if len(t) == 2 and t.upper() in ABBR_SET:
            return t.upper()
        abbr = STATE_TO_ABBR.get(re.sub(r'[^A-Za-z ]+', '', t).title())
        if abbr:
            return abbr
    letters = re.findall(r'[A-Za-z]', s)
    return ''.join(letters[:2]).upper() if letters else ""


def parse_qty(desc: str) -> int:
    if not isinstance(desc, str) or not desc.strip():
        return 0
    nums = [int(n) for n in re.findall(r'(\d+)\s*x', desc.lower())]
    if nums:
        return sum(nums)
    alt = [int(n) for n in re.findall(r'\b(\d+)\b', desc)]
    return sum(alt) if alt else 0


def split_qty(qty: int, chunk: int = 12):
    if qty <= 0:
        return [0]
    full, rem = divmod(qty, chunk)
    parts = [chunk] * full
    if rem:
        parts.append(rem)
    return parts or [0]


uploaded = st.file_uploader("📤 Carica il CSV", type=["csv"])

if uploaded is not None:
    # lettura CSV (tenta autodetect, poi ;)
    try:
        df = pd.read_csv(uploaded, dtype=str, sep=None, engine="python")
    except Exception:
        uploaded.seek(0)
        df = pd.read_csv(uploaded, dtype=str, sep=";", engine="python")
    df.columns = [c.strip() for c in df.columns]

    st.subheader("Anteprima CSV")
    st.dataframe(df.head(20))

    if st.button("👉 Genera XLSX"):
        # carico il template .xlsx dal repo
        template_df = pd.read_excel(TEMPLATE_PATH, dtype=str, engine="openpyxl")
        template_df.columns = [c.strip() for c in template_df.columns]
        template_columns = list(template_df.columns)
        template_cols_lower = {c.lower(): c for c in template_columns}

        # filtro solo USA
        df_usa = df[df["Paese"].fillna("").str.strip().str.casefold() == "united states of america"].copy()

        # mappo i campi del template
        mapping_candidates = {
            "OrderNo*": ["OrderNo*", "Order No", "OrderNo"],
            "Name*": ["Name*", "Name"],
            "Add1*": ["Add1*", "Address1", "Address 1", "Add 1"],
            "City*": ["City*", "City"],
            "State*": ["State*", "State"],
            "Zip*": ["Zip*", "ZIP", "Postal Code", "Postcode"],
            "Phone*": ["Phone*", "Phone", "Telephone"],
            "Ice Packs (Yes/No)*": ["Ice Packs (Yes/No)*", "Ice Packs"],
            "QTY*": ["QTY*", "Qty", "Quantity"],
            "Weight*": ["Weight*", "Weight"],
            "Email": ["Email", "E-mail", "Mail"],
            "SKU*": ["SKU*", "SKU"],
            "Insurance": ["Insurance", "Insurance Amount", "Insured Value"],
            "PC Type*": ["PC Type*", "PC Type"],
        }

        def resolve_col(key, candidates):
            for c in candidates:
                if c.lower() in template_cols_lower:
                    return template_cols_lower[c.lower()]
            # se non esiste nel template, la aggiungo in fondo
            if key not in template_columns:
                template_columns.append(key)
            return key

        resolved = {k: resolve_col(k, v) for k, v in mapping_candidates.items()}

        rows = []
        for _, r in df_usa.iterrows():
            order_no = (r.get("ID Spedizione", "") or "").strip()
            name = r.get("Destinatario", "")
            addr = r.get("Indirizzo", "")
            city = r.get("Città", "")
            state = normalize_state(r.get("Provincia", ""))
            zipc = r.get("CAP", "")
            phone = r.get("Telefono", "")
            email = r.get("e-mail", "")
            insurance = r.get("Importo Netto Assicurazione", "")
            qty_total = parse_qty(r.get("Descrizione merce", ""))

            for q in split_qty(qty_total, 12):
                weight = int(ceil(3.66 * q)) if q else 0
                row = {col: "" for col in template_columns}
                row[resolved["OrderNo*"]] = order_no
                row[resolved["Name*"]] = name
                row[resolved["Add1*"]] = addr
                row[resolved["City*"]] = city
                row[resolved["State*"]] = state
                row[resolved["Zip*"]] = zipc
                row[resolved["Phone*"]] = phone
                row[resolved["Ice Packs (Yes/No)*"]] = "NO"
                row[resolved["QTY*"]] = q
                row[resolved["Weight*"]] = weight
                row[resolved["Email"]] = email
                row[resolved["SKU*"]] = "STILL WINE"
                row[resolved["Insurance"]] = insurance
                row[resolved["PC Type*"]] = "WINE"
                rows.append(row)

        df_out = pd.DataFrame(rows, columns=template_columns)

        st.success(f"Generato file con {len(df_out)} righe (solo USA).")

        # crea XLSX in memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Sheet1")
        output.seek(0)

        st.download_button(
            "📥 Scarica XLSX",
            data=output,
            file_name="wine2weship_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Carica un CSV per iniziare.")
