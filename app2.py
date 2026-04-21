import io
import sys
import streamlit as st
from extractor import extract_double
from comparator import comparer
from generator import generer_excel, generer_rapport_txt
import zipfile
from email.message import EmailMessage
from datetime import datetime
import smtplib


# 1. CONFIGURATION HAUTE PERFORMANCE
st.set_page_config(
    page_title="DGIDocExtract",
    layout="wide",
)

# 2. DESIGN SYSTEM : "NEXUS ULTRA" (CSS Avancé)
st.markdown("""
    <style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap');

    /* Reset global */
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
        color: #1f2937;
    }

    /* Gradient Background Global */
    .stApp {
        background: radial-gradient(circle at top right, #f8fafc, #eff6ff);
    }

    /* Barre Latérale Flottante */
    [data-testid="stSidebar"] {
        background: rgba(15, 23, 42, 0.95);
        backdrop-filter: blur(10px);
        border-right: 1px solid rgba(255,255,255,0.1);
        padding-top: 2rem;
    }

    /* Cartes de Score Glassmorphism */
    div[data-testid="stMetric"] {
        background: rgba(255, 255, 255, 0.8);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(255, 255, 255, 0.5);
        border-radius: 20px;
        padding: 25px !important;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05);
        transition: transform 0.3s ease;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-5px);
        border: 1px solid #3b82f6;
    }

    /* Bouton d'Action "Call to Action" */
    .stButton>button {
        width: 100%;
        background: linear-gradient(90deg, #2563eb 0%, #7c3aed 100%);
        color: white;
        border: none;
        padding: 15px 30px;
        border-radius: 12px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 1px;
        box-shadow: 0 4px 14px 0 rgba(37, 99, 235, 0.39);
        transition: all 0.4s;
    }
    .stButton>button:hover {
        transform: scale(1.02);
        box-shadow: 0 6px 20px rgba(124, 58, 237, 0.45);
    }

    /* Animation du Status */
    .stStatusWidget {
        border-radius: 15px;
        border-left: 5px solid #2563eb;
    }

    /* Custom Header */
    .nexus-header {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 2.5rem;
        border-radius: 24px;
        margin-bottom: 2rem;
        color: white;
        position: relative;
        overflow: hidden;
    }
    .nexus-header::after {
        content: "";
        position: absolute;
        top: -50%; right: -10%;
        width: 300px; height: 300px;
        background: rgba(59, 130, 246, 0.1);
        border-radius: 50%;
    }
    </style>
    """, unsafe_allow_html=True)

def envoyer_email(model_name="glm",rapport):
    try:
        msg = EmailMessage()
        msg["From"] = st.secrets["email_envoyeur"]
        msg["To"] = st.secrets["email_destinataire"]
        msg["Subject"] = f"Extraction - {datetime.now().strftime('%H:%M:%S')}"
        msg.set_content(f"Le bouton Extraction a été cliqué à {datetime.now()} sur l'application DGIDocExtract avec le model {model_name}\n voici le rapport :\n {rapport}")
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(st.secrets["email_envoyeur"], st.secrets["mot_de_passe_app"])
            server.send_message(msg)
        
        return True
    except Exception as e:
        st.error(f"Erreur envoi email: {e}")
        return False

# 3. STRUCTURE SIDEBAR (L'expérience utilisateur commence ici)
with st.sidebar:
    st.markdown("<h2 style='color: #60a5fa;'> To Excel</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #94a3b8; font-size: 0.8rem;'>DGIDocExtract</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("### Document Source")
    image_file = st.file_uploader("", type=["jpg", "jpeg", "png", "webp"], label_visibility="collapsed")
    
    if image_file:
        with st.expander("👁️ Aperçu du document", expanded=False):
            st.image(image_file)
        st.markdown("---")
        st.caption(" DocExtract  —  Extraction par intelligence artificielle  —  Usage interne")

# 4. MAIN INTERFACE
        
st.markdown("""
    <div class="nexus-header">
        <h4 style='color: #60a5fa; margin: 0;'>DGIDocExtract</h4>
        <h1 style='color: white; margin: 0; font-size: 2.5rem;'>Papier vers Excel structure</h1>
        <p style='color: #94a3b8; margin-top: 10px;'> double extraction avec controle de coherence integre.</p>
    </div>
    """, unsafe_allow_html=True)

if not image_file:
    st.info(" 👈Bienvenue. Veuillez glisser un document financier dans le panneau latéral pour initier l'analyse.")
else:
    if st.button("Extraction"):
        tmp_path = None
        try:
            # 1. Configuration
            api_key = st.secrets["ZHIPUAI_API_KEY"]
            
            # 2. Fichier temporaire (portable)
            import tempfile
            import os
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{image_file.name}") as tmp_file:
                tmp_file.write(image_file.getbuffer())
                tmp_path = tmp_file.name
            
            # 3. Extraction
            with st.status("🚀 Extraction en cours...", expanded=True) as status:
                try:
                    status.write("🌐 Appel à l'API Gemini...")
                    extraction_a, extraction_b = extract_double(api_key, tmp_path, engine="glm")
                    status.write("✅ Extraction réussie")
                except Exception as api_error:
                    status.update(label="❌ Échec API", state="error")
                    raise Exception(f"API_ERROR:{str(api_error)}")
                

                
                status.write("🔍 Comparaison...")
                resultat = comparer(extraction_a, extraction_b)
                
                status.write("📊 Génération Excel...")
                wb = generer_excel(resultat, tmp_path)
                buf_excel = io.BytesIO()
                wb.save(buf_excel)
                buf_excel.seek(0)
                
                rapport_txt = generer_rapport_txt(resultat, tmp_path)
                envoyer_email(rapport = rapport_txt)
                
                status.update(label="✅ Terminé", state="complete")
            
            # 4. Affichage résultats
            rapport = resultat["rapport"]
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("INTÉGRITÉ", f"{rapport['score_global']}%")
            with col2:
                st.metric("CRITIQUES", rapport["rouges"], delta="Attention")
            with col3:
                st.metric("INCERTITUDES", rapport["oranges"])
            
            # 5. Exports
            nom_base = image_file.name.rsplit(".", 1)[0]
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr(f"{nom_base}_DATA.xlsx", buf_excel.getvalue())
                zip_file.writestr(f"{nom_base}_AUDIT_LOG.txt", rapport_txt.encode('utf-8'))
            zip_buffer.seek(0)
            
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("📦 ZIP COMPLET", zip_buffer, f"{nom_base}_DGI.zip", use_container_width=True)
            with c2:
                st.download_button("📊 EXCEL", buf_excel, f"{nom_base}.xlsx", use_container_width=True)
        
        except Exception as e:
            error_str = str(e)
            if "API_ERROR:" in error_str:
                st.error(f"🚨 **Erreur API Gemini**\n\n{error_str.replace('API_ERROR:', '')}")
            elif "429" in error_str or "quota" in error_str.lower():
                st.error("🚨 **Quota API atteint**\n\nRéessayez plus tard ou contactez l'administrateur.")
            else:
                st.error(f"🚨 **Erreur**\n\n```\n{error_str[:300]}\n```")
                # Option debug
                if st.checkbox("Afficher l'erreur complète"):
                    st.exception(e)
        
        finally:
            # Nettoyage
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except:
                    pass

        


