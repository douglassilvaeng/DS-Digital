import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime

try:
    from streamlit_js_eval import get_geolocation
except Exception:
    get_geolocation = None


# =====================================================
# CONFIGURAÇÕES
# =====================================================

st.set_page_config(
    page_title="Diálogo de Segurança Digital",
    page_icon="🛡️",
    layout="wide"
)

TEMA_DDS = "Cuidados com Animais Peçonhentos"
ARQUIVO_EXCEL = "dds_registros.xlsx"
ARQUIVO_PESSOAS = "pessoas.csv"
PASTA_IMAGENS = Path("imagens")
PASTA_ASSETS = Path("assets")
SENHA_ADMIN = "1234"


# =====================================================
# CSS
# =====================================================

def aplicar_css():
    st.markdown("""
    <style>
        .stApp {
            background-color: #f5f7fa;
        }

        .titulo-principal {
            background: linear-gradient(90deg, #003b2f, #007f5f);
            padding: 25px;
            border-radius: 18px;
            color: white;
            text-align: center;
            margin-bottom: 25px;
        }

        .card {
            background-color: white;
            padding: 22px;
            border-radius: 18px;
            box-shadow: 0px 4px 12px rgba(0,0,0,0.08);
            margin-bottom: 18px;
        }

        .card-alerta {
            background-color: #fff3cd;
            padding: 18px;
            border-left: 6px solid #ff9800;
            border-radius: 12px;
            margin-bottom: 18px;
        }
    </style>
    """, unsafe_allow_html=True)


# =====================================================
# BANCO DE DADOS
# =====================================================

def criar_excel_se_nao_existir():
    if not Path(ARQUIVO_EXCEL).exists():
        participantes = pd.DataFrame(columns=["Equipe", "Matrícula", "Nome"])

        presencas = pd.DataFrame(columns=[
            "Data", "Hora", "Nome", "Matrícula", "Equipe",
            "Tipo Participante", "Tema DDS", "Pontuação", "Status",
            "Declaração", "Latitude", "Longitude", "Link Localização"
        ])

        respostas = pd.DataFrame(columns=[
            "Data", "Hora", "Nome", "Matrícula", "Equipe",
            "Tipo Participante", "Pergunta", "Resposta", "Correta"
        ])

        with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
            participantes.to_excel(writer, sheet_name="Participantes", index=False)
            presencas.to_excel(writer, sheet_name="Presenças", index=False)
            respostas.to_excel(writer, sheet_name="Respostas", index=False)


def garantir_colunas(df, colunas):
    for coluna in colunas:
        if coluna not in df.columns:
            df[coluna] = ""
    return df[colunas]


def carregar_abas():
    criar_excel_se_nao_existir()

    participantes = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Participantes")
    presencas = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Presenças")
    respostas = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Respostas")

    participantes = garantir_colunas(participantes, ["Equipe", "Matrícula", "Nome"])

    presencas = garantir_colunas(presencas, [
        "Data", "Hora", "Nome", "Matrícula", "Equipe",
        "Tipo Participante", "Tema DDS", "Pontuação", "Status",
        "Declaração", "Latitude", "Longitude", "Link Localização"
    ])

    respostas = garantir_colunas(respostas, [
        "Data", "Hora", "Nome", "Matrícula", "Equipe",
        "Tipo Participante", "Pergunta", "Resposta", "Correta"
    ])

    participantes["Matrícula"] = participantes["Matrícula"].astype(str).str.strip()
    presencas["Matrícula"] = presencas["Matrícula"].astype(str).str.strip()
    respostas["Matrícula"] = respostas["Matrícula"].astype(str).str.strip()

    return participantes, presencas, respostas


def salvar_abas(participantes, presencas, respostas):
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
        participantes.to_excel(writer, sheet_name="Participantes", index=False)
        presencas.to_excel(writer, sheet_name="Presenças", index=False)
        respostas.to_excel(writer, sheet_name="Respostas", index=False)


# =====================================================
# IMPORTAR PARTICIPANTES
# =====================================================

def importar_pessoas_csv():
    if Path(ARQUIVO_PESSOAS).exists():
        df = pd.read_csv(ARQUIVO_PESSOAS, sep=None, engine="python")

        df = df.rename(columns={
            "des_equipe": "Equipe",
            "Matricula": "Matrícula",
            "Nome completo": "Nome"
        })

        df = df[["Equipe", "Matrícula", "Nome"]]

        df["Equipe"] = df["Equipe"].astype(str).str.strip()
        df["Matrícula"] = df["Matrícula"].astype(str).str.strip()
        df["Nome"] = df["Nome"].astype(str).str.strip()

        participantes, presencas, respostas = carregar_abas()
        participantes = df.drop_duplicates(subset=["Matrícula"])

        salvar_abas(participantes, presencas, respostas)

        return True, len(participantes)

    return False, 0


# =====================================================
# LOCALIZAÇÃO
# =====================================================

def obter_localizacao():
    latitude = ""
    longitude = ""
    link_maps = ""

    if get_geolocation is not None:
        try:
            location = get_geolocation()

            if location and "coords" in location:
                latitude = location["coords"].get("latitude", "")
                longitude = location["coords"].get("longitude", "")

                if latitude != "" and longitude != "":
                    link_maps = f"https://www.google.com/maps?q={latitude},{longitude}"

        except Exception:
            pass

    return latitude, longitude, link_maps


# =====================================================
# SALVAR REGISTRO
# =====================================================

def salvar_registro(
    nome,
    matricula,
    equipe,
    tipo_participante,
    pontuacao,
    respostas_usuario,
    latitude,
    longitude,
    link_maps
):
    participantes, presencas, respostas = carregar_abas()

    agora = datetime.now()
    data = agora.strftime("%d/%m/%Y")
    hora = agora.strftime("%H:%M:%S")

    nova_presenca = {
        "Data": data,
        "Hora": hora,
        "Nome": nome,
        "Matrícula": str(matricula),
        "Equipe": equipe,
        "Tipo Participante": tipo_participante,
        "Tema DDS": TEMA_DDS,
        "Pontuação": pontuacao,
        "Status": "Presente",
        "Declaração": "Confirmado",
        "Latitude": latitude,
        "Longitude": longitude,
        "Link Localização": link_maps
    }

    presencas = pd.concat(
        [presencas, pd.DataFrame([nova_presenca])],
        ignore_index=True
    )

    novas_respostas = []

    for item in respostas_usuario:
        novas_respostas.append({
            "Data": data,
            "Hora": hora,
            "Nome": nome,
            "Matrícula": str(matricula),
            "Equipe": equipe,
            "Tipo Participante": tipo_participante,
            "Pergunta": item["pergunta"],
            "Resposta": item["resposta"],
            "Correta": "Sim" if item["correta"] else "Não"
        })

    respostas = pd.concat(
        [respostas, pd.DataFrame(novas_respostas)],
        ignore_index=True
    )

    salvar_abas(participantes, presencas, respostas)


# =====================================================
# PERGUNTAS
# =====================================================

def obter_perguntas():
    return [
        {
            "pergunta": "Onde animais peçonhentos podem estar escondidos?",
            "opcoes": [
                "Apenas em florestas",
                "Caixas de medição, vegetação, entulhos e frestas",
                "Somente dentro de veículos"
            ],
            "correta": "Caixas de medição, vegetação, entulhos e frestas"
        },
        {
            "pergunta": "O que fazer ao encontrar uma cobra ou escorpião?",
            "opcoes": [
                "Tentar capturar o animal",
                "Afastar-se, manter distância e comunicar a equipe",
                "Continuar o serviço normalmente"
            ],
            "correta": "Afastar-se, manter distância e comunicar a equipe"
        },
        {
            "pergunta": "Qual EPI ajuda na proteção contra picadas nas pernas?",
            "opcoes": [
                "Perneira",
                "Protetor auricular",
                "Cinto de segurança"
            ],
            "correta": "Perneira"
        },
        {
            "pergunta": "Em caso de acidente com animal peçonhento, o correto é:",
            "opcoes": [
                "Fazer torniquete",
                "Cortar o local da picada",
                "Encaminhar rapidamente ao serviço de saúde"
            ],
            "correta": "Encaminhar rapidamente ao serviço de saúde"
        },
        {
            "pergunta": "Antes de iniciar o serviço, a equipe deve:",
            "opcoes": [
                "Inspecionar visualmente o ambiente",
                "Colocar a mão em buracos para verificar",
                "Ignorar vegetação e entulhos"
            ],
            "correta": "Inspecionar visualmente o ambiente"
        }
    ]


# =====================================================
# TELA PARTICIPANTE
# =====================================================

def tela_participante():
    st.markdown(f"""
    <div class="titulo-principal">
        <h1>🛡️ Diálogo de Segurança Digital</h1>
        <h3>{TEMA_DDS}</h3>
    </div>
    """, unsafe_allow_html=True)

    participantes, presencas, respostas = carregar_abas()

    st.info("Digite sua matrícula. Caso não esteja na lista, o acesso será registrado como convidado.")

    with st.form("form_matricula"):
        matricula_digitada = st.text_input("Matrícula")
        verificar = st.form_submit_button("Verificar matrícula")

    if verificar:
        if matricula_digitada.strip() == "":
            st.warning("Informe a matrícula.")
        else:
            matricula_digitada = matricula_digitada.strip()

            pessoa = participantes[
                participantes["Matrícula"].astype(str).str.strip() == matricula_digitada
            ]

            st.session_state["matricula"] = matricula_digitada

            if pessoa.empty:
                st.session_state["tipo_participante"] = "Convidado"
                st.session_state["matricula_encontrada"] = False
                st.session_state["liberado"] = False
                st.warning("Matrícula não encontrada. Continue como convidado.")
            else:
                st.session_state["tipo_participante"] = "Cadastrado"
                st.session_state["matricula_encontrada"] = True
                st.session_state["nome"] = pessoa.iloc[0]["Nome"]
                st.session_state["equipe"] = pessoa.iloc[0]["Equipe"]
                st.session_state["liberado"] = True
                st.success("Matrícula encontrada. Acesso liberado.")

    if st.session_state.get("matricula_encontrada") is False:
        st.subheader("Identificação do Convidado")

        with st.form("form_convidado"):
            nome_convidado = st.text_input("Nome completo")
            equipe_convidado = st.text_input("Equipe, empresa ou setor")
            confirmar_convidado = st.form_submit_button("Continuar como convidado")

        if confirmar_convidado:
            if nome_convidado.strip() == "" or equipe_convidado.strip() == "":
                st.warning("Informe nome e equipe/setor.")
            else:
                st.session_state["nome"] = nome_convidado.strip()
                st.session_state["equipe"] = equipe_convidado.strip()
                st.session_state["tipo_participante"] = "Convidado"
                st.session_state["liberado"] = True
                st.success("Acesso liberado como convidado.")

    if st.session_state.get("liberado"):
        nome = st.session_state["nome"]
        matricula = st.session_state["matricula"]
        equipe = st.session_state["equipe"]
        tipo_participante = st.session_state.get("tipo_participante", "Cadastrado")

        st.success(
            f"Participante: {nome} | Matrícula: {matricula} | "
            f"Equipe/Setor: {equipe} | Tipo: {tipo_participante}"
        )

        ja_participou = presencas[
            presencas["Matrícula"].astype(str).str.strip() == str(matricula).strip()
        ]

        if not ja_participou.empty:
            st.warning("Essa matrícula já possui presença registrada para este DDS.")

        st.divider()
        st.header("📍 Localização")

        st.write("Permita o acesso à localização quando o navegador solicitar.")

        latitude, longitude, link_maps = obter_localizacao()

        if latitude != "" and longitude != "":
            st.success("Localização capturada com sucesso.")
            st.write(f"Latitude: `{latitude}`")
            st.write(f"Longitude: `{longitude}`")
            st.write(f"[Abrir no Google Maps]({link_maps})")
        else:
            st.warning("Localização não capturada. O registro seguirá com data e hora.")

        st.divider()
        st.header("📚 Conteúdo do DDS")

        if PASTA_IMAGENS.exists():
            imagens = sorted(PASTA_IMAGENS.glob("*.png"))

            if imagens:
                for imagem in imagens:
                    st.image(str(imagem), use_container_width=True)
            else:
                st.warning("A pasta imagens está vazia.")
        else:
            st.warning("Crie a pasta imagens e coloque os slides em PNG.")

        st.divider()
        st.header("❓ Quiz de Verificação")

        perguntas = obter_perguntas()
        respostas_usuario = []

        with st.form("form_quiz"):
            for i, p in enumerate(perguntas, start=1):
                resposta = st.radio(
                    f"{i}. {p['pergunta']}",
                    p["opcoes"],
                    key=f"pergunta_{i}"
                )

                respostas_usuario.append({
                    "pergunta": p["pergunta"],
                    "resposta": resposta,
                    "correta": resposta == p["correta"]
                })

            declaracao = st.checkbox(
                "Declaro que participei do DDS, visualizei o conteúdo e respondi às perguntas de verificação."
            )

            finalizar = st.form_submit_button("Finalizar e registrar presença")

        if finalizar:
            if not declaracao:
                st.warning("Marque a declaração para registrar a presença.")
            elif not ja_participou.empty:
                st.error("Presença já registrada anteriormente.")
            else:
                acertos = sum(1 for r in respostas_usuario if r["correta"])
                pontuacao = f"{acertos}/{len(perguntas)}"

                salvar_registro(
                    nome,
                    matricula,
                    equipe,
                    tipo_participante,
                    pontuacao,
                    respostas_usuario,
                    latitude,
                    longitude,
                    link_maps
                )

                st.success("✅ Presença registrada com sucesso!")
                st.write(f"Pontuação: **{pontuacao}**")


# =====================================================
# TELA ADMIN
# =====================================================

def tela_admin():
    st.markdown("""
    <div class="titulo-principal">
        <h1>📊 Administração do DDS</h1>
        <h3>Painel de Controle</h3>
    </div>
    """, unsafe_allow_html=True)

    senha = st.text_input("Senha do administrador", type="password")

    if senha != SENHA_ADMIN:
        st.warning("Informe a senha para acessar.")
        return

    participantes, presencas, respostas = carregar_abas()

    st.success("Acesso liberado.")

    aba1, aba2, aba3, aba4 = st.tabs([
        "Importar participantes",
        "Resumo",
        "Presentes e Pendentes",
        "Respostas"
    ])

    with aba1:
        st.subheader("Importar pessoas.csv")

        st.write("O arquivo precisa estar na mesma pasta do app.py com o nome `pessoas.csv`.")

        if st.button("Importar lista de participantes"):
            ok, qtd = importar_pessoas_csv()

            if ok:
                st.success(f"Lista importada com sucesso. Total: {qtd} participantes.")
                st.rerun()
            else:
                st.error("Arquivo pessoas.csv não encontrado.")

        st.dataframe(participantes, use_container_width=True)

    with aba2:
        st.subheader("Resumo Geral")

        presencas_cadastrados = presencas[
            presencas["Tipo Participante"] == "Cadastrado"
        ]

        presencas_convidados = presencas[
            presencas["Tipo Participante"] == "Convidado"
        ]

        total_cadastrados = len(participantes)
        total_cadastrados_presentes = presencas_cadastrados["Matrícula"].astype(str).nunique()
        total_convidados = presencas_convidados["Matrícula"].astype(str).nunique()
        total_pendentes = total_cadastrados - total_cadastrados_presentes
        total_geral_presentes = presencas["Matrícula"].astype(str).nunique()

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Cadastrados", total_cadastrados)
        col2.metric("Cadastrados presentes", total_cadastrados_presentes)
        col3.metric("Convidados", total_convidados)
        col4.metric("Pendentes", total_pendentes)

        st.metric("Total geral de presenças", total_geral_presentes)

        if not participantes.empty:
            resumo_equipe = participantes.groupby("Equipe").size().reset_index(name="Cadastrados")

            presentes_equipe = presencas_cadastrados.groupby("Equipe")["Matrícula"].nunique().reset_index(name="Presentes")

            resumo_equipe = resumo_equipe.merge(
                presentes_equipe,
                on="Equipe",
                how="left"
            )

            resumo_equipe["Presentes"] = resumo_equipe["Presentes"].fillna(0).astype(int)
            resumo_equipe["Pendentes"] = resumo_equipe["Cadastrados"] - resumo_equipe["Presentes"]

            st.subheader("Resumo por Equipe")
            st.dataframe(resumo_equipe, use_container_width=True)

    with aba3:
        st.subheader("Presentes e Pendentes")

        matriculas_presentes_cadastrados = presencas[
            presencas["Tipo Participante"] == "Cadastrado"
        ]["Matrícula"].astype(str).unique()

        presentes_cadastrados = participantes[
            participantes["Matrícula"].astype(str).isin(matriculas_presentes_cadastrados)
        ]

        pendentes = participantes[
            ~participantes["Matrícula"].astype(str).isin(matriculas_presentes_cadastrados)
        ]

        convidados = presencas[
            presencas["Tipo Participante"] == "Convidado"
        ]

        st.write("### Cadastrados presentes")
        st.dataframe(presentes_cadastrados, use_container_width=True)

        st.write("### Pendentes")
        st.dataframe(pendentes, use_container_width=True)

        st.write("### Convidados presentes")
        st.dataframe(convidados, use_container_width=True)

        st.write("### Presenças completas")
        st.dataframe(presencas, use_container_width=True)

    with aba4:
        st.subheader("Respostas do Quiz")
        st.dataframe(respostas, use_container_width=True)

    st.divider()

    if Path(ARQUIVO_EXCEL).exists():
        with open(ARQUIVO_EXCEL, "rb") as arquivo:
            st.download_button(
                label="📥 Baixar relatório Excel",
                data=arquivo,
                file_name=ARQUIVO_EXCEL,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# =====================================================
# TELA SOBRE
# =====================================================

def tela_sobre():
    st.markdown('<div class="titulo-principal"><h1>🛡️ DDS Digital</h1><h3>Controle de Presença e Evidência de Participação</h3></div>', unsafe_allow_html=True)

    col1, col2 = st.columns([1, 2])

    with col1:
        imagem_sobre = PASTA_ASSETS / "Sobre.jpg"
        if imagem_sobre.exists():
            st.image(str(imagem_sobre), use_container_width=True)
        else:
            st.warning("Imagem não encontrada. Coloque em: assets/Sobre.jpg")

    with col2:
        # Removidas as quebras de linha excessivas dentro da string e ajustadas as tags
        html_sobre = """
        <div class="card">
            <h2>Sobre o Projeto</h2>
            <p>O <b>DDS Digital</b> foi desenvolvido para registrar a participação dos colaboradores em Diálogos Diários de Segurança, com controle de presença, quiz de verificação, data, hora e localização.</p>
            <p>O sistema identifica participantes cadastrados pela matrícula e também permite o registro de convidados.</p>
            <h3>Funcionalidades</h3>
            <ul>
                <li>Registro digital de presença</li>
                <li>Identificação por matrícula</li>
                <li>Acesso para convidados</li>
                <li>Quiz de verificação</li>
                <li>Registro de data, hora e localização</li>
                <li>Controle de presentes e pendentes</li>
                <li>Exportação em Excel</li>
            </ul>
            <p>Desenvolvido para apoiar a gestão de segurança, treinamentos e evidências operacionais das equipes de campo.</p>
        </div>
        """
        st.markdown(html_sobre, unsafe_allow_html=True)

    st.markdown('<div class="card-alerta"><b>Objetivo:</b> fortalecer a cultura de segurança e melhorar a evidência dos DDS realizados com as equipes de campo.</div>', unsafe_allow_html=True)


# =====================================================
# EXECUÇÃO
# =====================================================

criar_excel_se_nao_existir()
aplicar_css()

menu = st.sidebar.radio(
    "Menu",
    [
        "Participar do DDS",
        "Administração",
        "Sobre"
    ]
)

if menu == "Participar do DDS":
    tela_participante()
elif menu == "Administração":
    tela_admin()
else:
    tela_sobre()