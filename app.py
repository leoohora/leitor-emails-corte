
import streamlit as st
from imapclient import IMAPClient
import pyzmail

st.set_page_config(page_title="Leitor de Emails da Corte", page_icon="📨")
st.title("📨 Leitor de Emails da Corte de Imigração")

st.sidebar.header("🔑 Login no Outlook")

email = st.sidebar.text_input("Email Outlook", placeholder="seuemail@outlook.com")
senha = st.sidebar.text_input("Senha (ou senha de app)", type="password")
remetente = st.sidebar.text_input("Buscar por remetente", value="eoir@usdoj.gov")

if st.sidebar.button("📥 Ler Emails"):
    if not email or not senha:
        st.error("⚠️ Preencha seu email e senha.")
    else:
        try:
            with st.spinner("Conectando no Outlook..."):
                IMAP_SERVER = 'imap-mail.outlook.com'

                with IMAPClient(IMAP_SERVER) as client:
                    client.login(email, senha)
                    client.select_folder('INBOX', readonly=True)

                    st.success("✅ Conectado com sucesso!")

                    mensagens = client.search(['FROM', remetente])

                    st.info(f"🔍 Foram encontrados **{len(mensagens)}** emails desse remetente.")

                    if len(mensagens) == 0:
                        st.stop()

                    response = client.fetch(mensagens, ['ENVELOPE', 'BODY[]'])

                    for uid, data in response.items():
                        envelope = data[b'ENVELOPE']
                        assunto = envelope.subject.decode() if envelope.subject else "(sem assunto)"
                        data_email = envelope.date

                        st.subheader(f"✉️ {assunto}")
                        st.caption(f"📅 {data_email}")

                        message = pyzmail36.PyzMessage.factory(data[b'BODY[]'])

                        if message.text_part:
                            texto = message.text_part.get_payload().decode(message.text_part.charset)
                            with st.expander("📃 Ver corpo do email"):
                                st.write(texto)

                        if message.mailparts:
                            for part in message.mailparts:
                                if part.is_attachment():
                                    nome_arquivo = part.filename
                                    conteudo = part.get_payload()

                                    st.download_button(
                                        label=f"📄 Baixar {nome_arquivo}",
                                        data=conteudo,
                                        file_name=nome_arquivo
                                    )
                        st.divider()
        except Exception as e:
            st.error(f"❌ Erro: {str(e)}")
