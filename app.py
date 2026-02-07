# -*- coding: utf-8 -*-
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from docxtpl import DocxTemplate, RichText
import io
from datetime import datetime

# --- НАЛАШТУВАННЯ ---
st.set_page_config(page_title="Акт СПЗ", page_icon="🔥", layout="centered")
st.markdown("<style>.stButton button {width: 100%; background-color: #28a745; color: white;}</style>", unsafe_allow_html=True)

# --- ДОПОМІЖНІ ФУНКЦІЇ ---
def find_worksheet_case_insensitive(sh, name):
    try:
        return sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        for ws in sh.worksheets():
            if ws.title.lower() == name.lower():
                return ws
        return None

# --- ЗАВАНТАЖЕННЯ ДАНИХ ---
@st.cache_data(ttl=60)
def get_gsheet_data():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)
        sh = client.open_by_key(st.secrets["spreadsheet_id"])
        
        data = {}
        # 1. Системи
        ws_gen = find_worksheet_case_insensitive(sh, "загальні дані")
        if not ws_gen: st.error("Не знайдено вкладку 'загальні дані'"); return None
        data['systems'] = {r['Назва']: r['Код'] for r in ws_gen.get_all_records() if r['Код']}
        
        # 2. Ліцензіати
        ws_lic = find_worksheet_case_insensitive(sh, "Ліцензіати")
        lic_rows = ws_lic.get_all_records() if ws_lic else []
        data['licensees'] = {r['Short Name']: r['Full Text'].strip() for r in lic_rows if r['Short Name']}
        
        # 3. Підписанти
        ws_sig = find_worksheet_case_insensitive(sh, "Підписанти")
        raw_sigs = ws_sig.get_all_records() if ws_sig else []
        for p in raw_sigs:
            if not p.get('Label'): 
                p['Label'] = p.get('Name', 'Невідомо')
        data['signatories'] = raw_sigs

        # 4. Дефекти
        data['defects'] = {}
        for sys_name, sys_code in data['systems'].items():
            ws_sys = find_worksheet_case_insensitive(sh, sys_code)
            if ws_sys:
                recs = ws_sys.get_all_records()
                sys_defects = []
                for r in recs:
                    if r.get('Full Text'):
                        lbl = f"[{r.get('Category','?')}] {r.get('Short Name','?')}"
                        sys_defects.append({'label': lbl, 'full': r['Full Text']})
                data['defects'][sys_code] = sys_defects
            else:
                data['defects'][sys_code] = []
        return data
    except Exception as e:
        st.error(f"Помилка з'єднання: {e}")
        return None

# --- ГОЛОВНА ЛОГІКА ---
def main():
    st.title("🔥 Акт Невідповідності")
    
    keys_to_init = ['inst_pos', 'inst_name', 'maint_pos', 'maint_name', 'obs_pos', 'obs_name']
    for k in keys_to_init:
        if k not in st.session_state: st.session_state[k] = ""

    data_dict = get_gsheet_data()
    if not data_dict: return

    # 1. ОБ'ЄКТ
    with st.expander("🏢 1. Дані об'єкта", expanded=True):
        legal_name = st.text_input("Власник", placeholder="ТОВ...")
        legal_addr = st.text_input("Юр. адреса")
        c1, c2 = st.columns(2)
        obj_name = c1.text_input("Назва об'єкта")
        obj_addr = c2.text_input("Адреса об'єкта")
        project_info = st.text_area("Проектні дані", height=70)
        
        lic_opts = ["Ввести вручну..."] + list(data_dict['licensees'].keys())
        sel_lic = st.selectbox("Ліцензіат (Монтажна орг.)", lic_opts, index=0)
        
        if sel_lic == "Ввести вручну...":
            license_text = st.text_area("Текст ліцензії (введіть свій варіант)")
        else:
            license_text = st.text_area("Текст ліцензії", value=data_dict['licensees'][sel_lic])

    # 2. СИСТЕМИ
    st.subheader("🛠 2. Системи")
    sys_map = data_dict['systems']
    selected_sys = st.multiselect("Оберіть системи:", list(sys_map.keys()), default=list(sys_map.keys()))
    
    results_rt = {} 
    
    for sys_name, code in sys_map.items():
        if sys_name in selected_sys:
            defects = data_dict['defects'].get(code, [])
            opts_map = {d['label']: d['full'] for d in defects}
            
            with st.expander(f"{sys_name}", expanded=False):
                picked = st.multiselect(f"Порушення ({code})", list(opts_map.keys()))
                custom = st.text_area(f"Свій текст ({code}) - кожне зауваження з нового рядка", height=68)
                
                full_texts = [opts_map[p] for p in picked]
                if custom:
                    for line in custom.split('\n'):
                        if line.strip(): full_texts.append(line.strip())
                
                if full_texts:
                    txt = "".join([f"{i}. {t}\n" for i, t in enumerate(full_texts, 1)])
                    results_rt[code] = RichText(txt.strip())
                else:
                    results_rt[code] = "—"
        else:
            results_rt[code] = "—"

    # 3. КОМІСІЯ
    st.subheader("✍️ 3. Комісія")
    
    def update_person_fields(key_prefix, people_list):
        selected_label = st.session_state[f"{key_prefix}_sel"]
        if selected_label != "Ввести вручну...":
            p_data = next((p for p in people_list if str(p['Label']) == selected_label), None)
            if p_data:
                st.session_state[f"{key_prefix}_pos"] = p_data.get('Position', '')
                st.session_state[f"{key_prefix}_name"] = p_data.get('Name', '')

    def hybrid_selector_label(label, category, key_prefix):
        people = [s for s in data_dict['signatories'] if str(s.get('Category', '')).strip().lower() == category.lower()]
        opts = ["Ввести вручну..."] + [str(p['Label']) for p in people]
        
        st.selectbox(f"Оберіть зі списку ({label})", opts, key=f"{key_prefix}_sel", on_change=update_person_fields, args=(key_prefix, people))
        st.text_input(f"Посада ({label})", key=f"{key_prefix}_pos")
        st.text_input(f"ПІБ ({label})", key=f"{key_prefix}_name")

    c1, c2 = st.columns(2)
    with c1:
        cm_pos = st.text_input("Посада (Зам)", "Директор")
        cm_name = st.text_input("ПІБ (Зам)")
    with c2:
        cr_pos = st.text_input("Посада (Відп)", "Відповідальний за ПБ")
        cr_name = st.text_input("ПІБ (Відп)")

    st.markdown("---")
    col_i, col_m, col_o = st.columns(3)
    with col_i: hybrid_selector_label("Монтажник", "Installer", "inst")
    with col_m: hybrid_selector_label("ТО", "Maintenance", "maint")
    with col_o: hybrid_selector_label("Спостерігання", "Observer", "obs")

    st.markdown("---")
    
    dsns_people = [s for s in data_dict['signatories'] if str(s.get('Category','')).strip().upper() == 'DSNS']
    dsns_map = {str(p['Label']): p for p in dsns_people}
    
    sel_dsns_labels = st.multiselect("ДСНС (макс 3) - пошук за прізвищем", list(dsns_map.keys()), max_selections=3)

    if st.button("📝 СФОРМУВАТИ АКТ"):
        if not obj_name: st.error("Введіть назву об'єкта!"); return

        context = {
            'LEGAL': legal_name, 'LEGAL_ADDR': legal_addr, 'OBJECT': obj_name, 'ADDRESS': obj_addr,
            'PROJECT': project_info, 'LICENSE': license_text,
            'CLIENT_MAIN_POS': cm_pos, 'CLIENT_MAIN_NAME': cm_name,
            'CLIENT_RESP_POS': cr_pos, 'CLIENT_RESP_NAME': cr_name,
            'INSTALLER_POS': st.session_state['inst_pos'], 'INSTALLER_NAME': st.session_state['inst_name'],
            'MAINTENANCE_POS': st.session_state['maint_pos'], 'MAINTENANCE_NAME': st.session_state['maint_name'],
            'OBSERVER_POS': st.session_state['obs_pos'], 'OBSERVER_NAME': st.session_state['obs_name'],
        }
        context.update(results_rt)

        # --- НОВА ЛОГІКА ДСНС (Список) ---
        dsns_list = []
        for lbl in sel_dsns_labels:
            p = dsns_map.get(lbl)
            if p:
                dsns_list.append({'pos': p.get('Position', ''), 'name': p.get('Name', '')})
        
        # Передаємо список у шаблон
        context['dsns_list'] = dsns_list

        try:
            doc = DocxTemplate("template.docx")
            doc.render(context)
            buf = io.BytesIO(); doc.save(buf); buf.seek(0)
            
            st.success("Документ готовий!")
            st.download_button("⬇️ ЗАВАНТАЖИТИ DOCX", buf, f"Act_{datetime.now().strftime('%Y-%m-%d')}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Помилка: {e}")

if __name__ == "__main__":
    main()