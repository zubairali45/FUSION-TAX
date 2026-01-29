import streamlit as st

def show():
    # Tombol kembali ke beranda
    if st.button("← Kembali ke Beranda Utama"):
        st.session_state.current_page = 'beranda'
        st.session_state.selected_menu = None
        st.rerun()
    
    st.title("📋 DASHBOARD SISTEM PPPK")
    st.markdown("---")
    
    # Pilihan fitur/fungsi PPPK
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 🔎")
        st.markdown("### Croscheck Data")
        if st.button("Buka Croscheck PPPK", key="croscheck_pppk"):
            st.session_state.current_page = 'croscheck_pppk'
            st.session_state.selected_menu = 'croscheck'
            st.rerun()
        st.caption("Verifikasi data PPPK")
    
    with col2:
        st.markdown("### 💰")
        st.markdown("### Upload Pajak Gaji")
        if st.button("Upload Pajak Gaji", key="upload_pajak_pppk"):
            st.session_state.current_page = 'upload_pajak_gaji_pppk'
            st.rerun()
        st.caption("Upload data pajak gaji PPPK")
    
    with col3:
        st.markdown("### 🍽️")
        st.markdown("### Upload Pajak Makan")
        if st.button("Upload Pajak Makan", key="upload_makan_pppk"):
            st.session_state.current_page = 'upload_pajak_makan_pppk'
            st.rerun()
        st.caption("Upload data pajak makan PPPK")
    
    st.markdown("---")