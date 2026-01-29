import streamlit as st
import sys

# Konfigurasi halaman
st.set_page_config(
    page_title="FUSION-TAX - Beranda",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS untuk sembunyikan sidebar
st.markdown("""
    <style>
    [data-testid="stSidebar"] { display: none; }
    .stButton button { width: 100%; height: 120px; font-size: 20px !important; }
    .dashboard-card { 
        padding: 30px; 
        border-radius: 10px; 
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: transform 0.3s;
    }
    .dashboard-card:hover {
        transform: translateY(-5px);
    }
    </style>
""", unsafe_allow_html=True)

# Session state untuk navigasi
if 'current_page' not in st.session_state:
    st.session_state.current_page = 'beranda'
if 'selected_menu' not in st.session_state:
    st.session_state.selected_menu = None

# Fungsi navigasi
def navigate_to(page, menu=None):
    st.session_state.current_page = page
    if menu:
        st.session_state.selected_menu = menu
    st.rerun()

# BERANDA UTAMA (hanya 2 tombol)
if st.session_state.current_page == 'beranda':
    st.title("🏢 FUSION-TAX")
    st.markdown("---")
    
    col1, col2 = st.columns(2, gap="large")
    
    with col1:
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        st.markdown("### 👤")
        st.markdown("### SISTEM PNS")
        if st.button("MASUK DASHBOARD PNS", key="btn_pns"):
            navigate_to('dashboard_pns')
        st.markdown('</div>', unsafe_allow_html=True)
        
    with col2:
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        st.markdown("### 👥")
        st.markdown("### SISTEM PPPK")
        if st.button("MASUK DASHBOARD PPPK", key="btn_pppk"):
            navigate_to('dashboard_pppk')
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("*Kerja Praktek Universitas Muhammadiyah Pontianak*")
    st.markdown("*by Zubair Ali*")

# ROUTING ke dashboard dengan error handling
elif st.session_state.current_page == 'dashboard_pns':
    try:
        import dashboard_pns
        dashboard_pns.show()
    except Exception as e:
        st.error(f"Error loading PNS dashboard: {e}")
        if st.button("← Kembali ke Beranda Utama"):
            st.session_state.current_page = 'beranda'
            st.rerun()
    
elif st.session_state.current_page == 'dashboard_pppk':
    try:
        import dashboard_pppk
        dashboard_pppk.show()
    except Exception as e:
        st.error(f"Error loading PPPK dashboard: {e}")
        if st.button("← Kembali ke Beranda Utama"):
            st.session_state.current_page = 'beranda'
            st.rerun()

# ROUTING ke fitur PNS
elif st.session_state.current_page == 'croscheck_pns':
    try:
        import croscheck_pns
        croscheck_pns.show()
    except Exception as e:
        st.error(f"Error loading PNS croscheck: {e}")
        if st.button("← Kembali ke Beranda Utama"):
            st.session_state.current_page = 'beranda'
            st.rerun()

elif st.session_state.current_page == 'upload_pajak_gaji_pns':
    try:
        import upload_pajak_gaji_pns
        upload_pajak_gaji_pns.show()
    except Exception as e:
        st.error(f"Error loading PNS tax upload: {e}")
        if st.button("← Kembali ke Beranda Utama"):
            st.session_state.current_page = 'beranda'
            st.rerun()

elif st.session_state.current_page == 'upload_pajak_makan_pns':
    try:
        import upload_pajak_makan_pns
        upload_pajak_makan_pns.show()
    except Exception as e:
        st.error(f"Error loading PNS meal tax upload: {e}")
        if st.button("← Kembali ke Beranda Utama"):
            st.session_state.current_page = 'beranda'
            st.rerun()

# ROUTING ke fitur PPPK
elif st.session_state.current_page == 'croscheck_pppk':
    try:
        import croscheck_pppk
        croscheck_pppk.show()
    except SyntaxError as e:
        st.error(f"Syntax Error in croscheck_pppk.py: {e}")
        st.info("Silakan perbaiki syntax error di file croscheck_pppk.py terlebih dahulu.")
        if st.button("← Kembali ke Dashboard PPPK"):
            st.session_state.current_page = 'dashboard_pppk'
            st.rerun()
    except Exception as e:
        st.error(f"Error loading PPPK croscheck: {e}")
        if st.button("← Kembali ke Beranda Utama"):
            st.session_state.current_page = 'beranda'
            st.rerun()

elif st.session_state.current_page == 'upload_pajak_gaji_pppk':
    try:
        import upload_pajak_gaji_pppk
        upload_pajak_gaji_pppk.show()
    except Exception as e:
        st.error(f"Error loading PPPK tax upload: {e}")
        if st.button("← Kembali ke Beranda Utama"):
            st.session_state.current_page = 'beranda'
            st.rerun()

elif st.session_state.current_page == 'upload_pajak_makan_pppk':
    try:
        import upload_pajak_makan_pppk
        upload_pajak_makan_pppk.show()
    except Exception as e:
        st.error(f"Error loading PPPK meal tax upload: {e}")
        if st.button("← Kembali ke Beranda Utama"):
            st.session_state.current_page = 'beranda'
            st.rerun()
    
# ROUTING ke fitur PNS
elif st.session_state.current_page == 'upload_pajak_lembur_pns':
    try:
        import upload_pajak_lembur_pns
        upload_pajak_lembur_pns.show()
    except Exception as e:
        st.error(f"Error loading PNS overtime tax upload: {e}")
        if st.button("← Kembali ke Beranda Utama"):
            st.session_state.current_page = 'beranda'
            st.rerun()

# Fallback jika halaman tidak ditemukan
else:
    st.error("Halaman tidak ditemukan!")
    if st.button("🏠 Kembali ke Beranda Utama"):
        st.session_state.current_page = 'beranda'
        st.session_state.selected_menu = None
        st.rerun()