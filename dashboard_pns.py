import streamlit as st

def show():
    # Tombol kembali ke beranda
    if st.button("← Kembali ke Beranda Utama"):
        st.session_state.current_page = 'beranda'
        st.session_state.selected_menu = None
        st.rerun()
    
    st.title("📊 DASHBOARD SISTEM PNS")
    st.markdown("---")
    
    # Pilihan fitur/fungsi PNS - Baris pertama
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 🔍")
        st.markdown("### Croscheck Data Gaji")
        if st.button("Buka Croscheck PNS", key="croscheck_pns"):
            st.session_state.current_page = 'croscheck_pns'
            st.session_state.selected_menu = 'croscheck'
            st.rerun()
        st.caption("Verifikasi dan validasi data PNS")
    
    with col2:
        st.markdown("### 💰")
        st.markdown("### Upload Pajak Gaji")
        if st.button("Upload Pajak Gaji PNS", key="upload_pajak_gaji"):
            st.session_state.current_page = 'upload_pajak_gaji_pns'
            st.rerun()
        st.caption("Upload data pajak gaji PNS")
    
    with col3:
        st.markdown("### 🍽️")
        st.markdown("### Upload Pajak Makan")
        if st.button("Upload Pajak Makan PNS", key="upload_pajak_makan"):
            st.session_state.current_page = 'upload_pajak_makan_pns'
            st.rerun()
        st.caption("Upload data pajak dan tunjangan makan")
    
    # Baris kedua - Button Upload Pajak Lembur
    st.markdown("<br>", unsafe_allow_html=True)
    col4, col5, col6 = st.columns(3)
    
    with col4:
        st.markdown("### ⏰")
        st.markdown("### Upload Pajak Lembur")
        if st.button("Upload Pajak Lembur PNS", key="upload_pajak_lembur"):
            st.session_state.current_page = 'upload_pajak_lembur_pns'
            st.rerun()
        st.caption("Upload data pajak lembur PNS")
    
    st.markdown("---")