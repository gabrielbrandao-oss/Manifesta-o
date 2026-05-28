import streamlit as st
import pandas as pd
import numpy as np
import io
import requests
import re
import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from datetime import datetime

# ═══════════════════════════════════════════════════════════
# REGRAS DE NEGÓCIO
# ═══════════════════════════════════════════════════════════
SITUACOES_VALIDAS  = {'não informado', 'uso autorizado', 'doc não gerado'}
TERMOS_EXCLUSAO    = r'comodato|retorno|devolu[çc][aã]o'
COLUNAS_EXCLUSAO   = ['Observação', 'Natureza Operação', 'Tipo Documento SAP']
MAX_DOWNLOAD_MB    = 50
TIMEOUT_S          = 15
LIXO               = {'nan', 'none', '', 'nat', '<na>', 'null'}
ABA_CONTROLE       = 'Controle'

# Colunas visíveis na tabela principal (nesta ordem)
COLUNAS_EXIBIR = [
    'Nº NF',
    'Chave de Acesso',
    'Razão Social',
    'Valor Nota Fiscal',
    'Situação Manifestação',
]

# ═══════════════════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════════════════
st.set_page_config(page_title="Manifestações NF-e", layout="wide")
st.title("📋 Gestão de Manifestações NF-e")


# ═══════════════════════════════════════════════════════════
# UTILITÁRIOS
# ═══════════════════════════════════════════════════════════
def normalizar(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.upper()

def eh_lixo(s: pd.Series) -> pd.Series:
    return s.str.lower().isin(LIXO) | (s.str.len() <= 5)

def limpar_display(df: pd.DataFrame) -> pd.DataFrame:
    return df.astype(object).replace(
        {np.nan: '', pd.NaT: '', 'None': '', 'nan': '', '<NA>': '', 'NaT': ''}
    )

def validar_colunas(df, obrigatorias, aba):
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        st.error(f"Aba **{aba}** — colunas não encontradas: `{'`, `'.join(faltando)}`")
        st.stop()

def btn_exportar(df: pd.DataFrame, nome: str):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, index=False)
    st.download_button(
        f"⬇️ Exportar {nome}",
        data=buf.getvalue(),
        file_name=f"{nome.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ═══════════════════════════════════════════════════════════
# GOOGLE SHEETS — leitura e escrita
# ═══════════════════════════════════════════════════════════
def get_sheet_id() -> str:
    try:
        url = st.secrets["google_sheets"]["url"]
    except KeyError:
        st.error("secrets.toml sem a chave [google_sheets] url.")
        st.stop()
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    if not m:
        st.error("URL do Google Sheets inválida.")
        st.stop()
    return m.group(1)

def get_credentials():
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.readonly',
    ]
    if "gcp_service_account" not in st.secrets:
        st.error("Credenciais GCP não encontradas em secrets.toml. Necessário para salvar dados.")
        st.stop()
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    return creds

@st.cache_data(show_spinner=False, ttl=30)
def buscar_planilha_bytes() -> bytes:
    sheet_id   = get_sheet_id()
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    creds      = get_credentials()
    creds.refresh(Request())
    headers = {'Authorization': f'Bearer {creds.token}'}
    try:
        resp = requests.get(export_url, headers=headers, timeout=TIMEOUT_S)
        resp.raise_for_status()
        mb = len(resp.content) / (1024 * 1024)
        if mb > MAX_DOWNLOAD_MB:
            st.error(f"Arquivo muito grande ({mb:.1f} MB).")
            st.stop()
        return resp.content
    except requests.exceptions.HTTPError as e:
        st.error(f"Erro HTTP ao baixar planilha: {e}")
        st.stop()
    except requests.exceptions.RequestException as e:
        st.error(f"Falha de rede: {e}")
        st.stop()

@st.cache_data(show_spinner=False, ttl=30)
def carregar_controle() -> pd.DataFrame:
    """Lê a aba Controle da planilha (marcações e observações salvas)."""
    try:
        sheet_id = get_sheet_id()
        creds    = get_credentials()
        gc       = gspread.authorize(creds)
        sh       = gc.open_by_key(sheet_id)
        try:
            ws = sh.worksheet(ABA_CONTROLE)
            dados = ws.get_all_records()
            if dados:
                return pd.DataFrame(dados)
        except gspread.exceptions.WorksheetNotFound:
            # Cria a aba se não existir
            ws = sh.add_worksheet(title=ABA_CONTROLE, rows=5000, cols=5)
            ws.append_row(['Chave', 'Tratado', 'Observacao_Controle', 'Responsavel', 'Atualizado_em'])
    except Exception as e:
        st.warning(f"Não foi possível carregar controle salvo: {e}")
    return pd.DataFrame(columns=['Chave', 'Tratado', 'Observacao_Controle', 'Responsavel', 'Atualizado_em'])

def salvar_controle(df_controle: pd.DataFrame):
    """Grava o DataFrame de controle inteiro na aba Controle."""
    try:
        sheet_id = get_sheet_id()
        creds    = get_credentials()
        gc       = gspread.authorize(creds)
        sh       = gc.open_by_key(sheet_id)
        try:
            ws = sh.worksheet(ABA_CONTROLE)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=ABA_CONTROLE, rows=5000, cols=5)

        ws.clear()
        df_salvar = df_controle.fillna('').astype(str)
        ws.update([df_salvar.columns.tolist()] + df_salvar.values.tolist())
        # Limpa cache para recarregar na próxima vez
        carregar_controle.clear()
        st.success("✅ Alterações salvas com sucesso!")
    except Exception as e:
        st.error(f"Erro ao salvar: {e}")


# ═══════════════════════════════════════════════════════════
# PROCESSAMENTO FISCAL
# ═══════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def processar(file_bytes: bytes):
    df_sap  = pd.read_excel(io.BytesIO(file_bytes), sheet_name='SAP',  dtype=str, engine='openpyxl')
    df_qive = pd.read_excel(io.BytesIO(file_bytes), sheet_name='Qive', dtype=str, engine='openpyxl')

    validar_colunas(df_sap,  ['Chave de Acesso', 'Situação Manifestação', 'Tipo Doc. Fiscal'], 'SAP')
    validar_colunas(df_qive, ['Chave de acesso'], 'Qive')

    df_sap['Chave de Acesso']  = normalizar(df_sap['Chave de Acesso'])
    df_qive['Chave de acesso'] = normalizar(df_qive['Chave de acesso'])
    df_sap  = df_sap[~eh_lixo(df_sap['Chave de Acesso'])].copy()
    df_qive = df_qive[~eh_lixo(df_qive['Chave de acesso'])].copy()

    # ── Filtro base: NF-e E status válido ──
    mask_nfe = df_sap['Tipo Doc. Fiscal'].astype(str).str.strip().str.lower() == 'nf-e'
    mask_sit = df_sap['Situação Manifestação'].astype(str).str.strip().str.lower().isin(SITUACOES_VALIDAS)
    df_no_escopo   = df_sap[mask_nfe & mask_sit].copy()
    df_fora_escopo = df_sap[~(mask_nfe & mask_sit)].copy()

    # ── Exclusão por palavras nas colunas definidas ──
    mask_excl = pd.Series(False, index=df_no_escopo.index)
    for col in COLUNAS_EXCLUSAO:
        if col in df_no_escopo.columns:
            mask_excl |= df_no_escopo[col].astype(str).str.contains(TERMOS_EXCLUSAO, case=False, na=False)
    df_excluidos = df_no_escopo[mask_excl].copy()
    df_ativo     = df_no_escopo[~mask_excl].copy()

    # ── Cruzamento SAP × Qive ──
    df_merge = pd.merge(
        df_ativo, df_qive,
        left_on='Chave de Acesso', right_on='Chave de acesso',
        how='outer', indicator=True
    )

    # Emitente unificado — usa Razão Social do SAP, cai para Qive
    em_sap  = df_merge.get('Razão Social', pd.Series(dtype=str))
    em_qive = df_merge.get('Nome/Razão Social Emitente', pd.Series(dtype=str))
    df_merge['Razão Social'] = em_sap.replace(['nan','None',np.nan,''], pd.NA).fillna(em_qive)

    # Chave unificada
    ch_sap  = df_merge.get('Chave de Acesso', pd.Series(dtype=str))
    ch_qive = df_merge.get('Chave de acesso', pd.Series(dtype=str))
    df_merge['Chave de Acesso'] = ch_sap.replace(['nan','None',np.nan,''], pd.NA).fillna(ch_qive)

    df_merge['Status Cruzamento'] = df_merge['_merge'].map({
        'left_only':  'Lançado — falta na Qive',
        'right_only': 'Não gerado no SAP',
        'both':       'OK (ambos)',
    })

    # ── Pendências: apenas SAP, sem exclusões ──
    df_sap_lado = df_merge[df_merge['_merge'] != 'right_only'].copy()
    mask_pend   = df_sap_lado['Situação Manifestação'].astype(str).str.contains('não informado', case=False, na=False)
    if 'Manifestação' in df_sap_lado.columns:
        mask_pend |= df_sap_lado['Manifestação'].astype(str).str.contains('sem manifestação', case=False, na=False)

    df_pend = df_sap_lado[mask_pend].copy()

    # Barreira final nas pendências
    mask_excl_pend = pd.Series(False, index=df_pend.index)
    for col in COLUNAS_EXCLUSAO:
        if col in df_pend.columns:
            mask_excl_pend |= df_pend[col].astype(str).str.contains(TERMOS_EXCLUSAO, case=False, na=False)
    df_pend = df_pend[~mask_excl_pend].copy()

    return (
        limpar_display(df_merge),
        limpar_display(df_pend),
        limpar_display(df_fora_escopo),
        limpar_display(df_excluidos),
    )


# ═══════════════════════════════════════════════════════════
# MONTAR TABELA COM CONTROLE SALVO
# ═══════════════════════════════════════════════════════════
def montar_tabela_editavel(df: pd.DataFrame, df_controle: pd.DataFrame) -> pd.DataFrame:
    """
    Seleciona colunas visíveis, injeta Tratado e Observacao_Controle
    vindos da aba Controle (estado compartilhado salvo).
    """
    # Seleciona só as colunas que existem
    cols = [c for c in COLUNAS_EXIBIR if c in df.columns]
    df_view = df[cols].copy()

    # Garante coluna de chave para join com controle
    if 'Chave de Acesso' not in df_view.columns and 'Chave de Acesso' in df.columns:
        df_view['Chave de Acesso'] = df['Chave de Acesso'].values

    chave_col = 'Chave de Acesso'

    if not df_controle.empty and chave_col in df_view.columns:
        df_ctrl = df_controle[['Chave', 'Tratado', 'Observacao_Controle']].copy()
        df_ctrl.rename(columns={'Chave': chave_col}, inplace=True)
        df_view = df_view.merge(df_ctrl, on=chave_col, how='left')
        df_view['Tratado']            = df_view['Tratado'].fillna('').map(lambda x: x == 'True' or x is True)
        df_view['Observacao_Controle'] = df_view['Observacao_Controle'].fillna('')
    else:
        df_view['Tratado']            = False
        df_view['Observacao_Controle'] = ''

    # Reordena: Tratado primeiro, Observacao_Controle no fim
    col_order = ['Tratado'] + [c for c in df_view.columns if c not in ('Tratado', 'Observacao_Controle')] + ['Observacao_Controle']
    df_view = df_view[[c for c in col_order if c in df_view.columns]]

    return df_view


# ═══════════════════════════════════════════════════════════
# FILTROS INLINE
# ═══════════════════════════════════════════════════════════
def filtros(df: pd.DataFrame, key: str) -> pd.DataFrame:
    c1, c2, c3 = st.columns([3, 2, 1])
    with c1:
        busca = st.text_input("busca", key=f"busca_{key}",
                              placeholder="Buscar por Nº NF, emitente ou chave…",
                              label_visibility="collapsed")
    with c2:
        opts = ["Todos"]
        if 'Situação Manifestação' in df.columns:
            opts += sorted(df['Situação Manifestação'].dropna().unique().tolist())
        sel = st.selectbox("status", opts, key=f"status_{key}", label_visibility="collapsed")
    with c3:
        if st.button("↺", key=f"reset_{key}", help="Limpar filtros"):
            for k in [f"busca_{key}", f"status_{key}"]:
                st.session_state.pop(k, None)
            st.rerun()

    res = df.copy()
    if busca:
        mask = pd.Series(False, index=res.index)
        for col in ['Nº NF', 'Chave de Acesso', 'Razão Social']:
            if col in res.columns:
                mask |= res[col].astype(str).str.contains(busca, case=False, na=False)
        res = res[mask]
    if sel != "Todos" and 'Situação Manifestação' in res.columns:
        res = res[res['Situação Manifestação'] == sel]
    return res


# ═══════════════════════════════════════════════════════════
# RENDERIZA TABELA EDITÁVEL + BOTÃO SALVAR
# ═══════════════════════════════════════════════════════════
def render_editavel(df_dados: pd.DataFrame, df_controle: pd.DataFrame,
                    key: str, nome_export: str, responsavel: str):
    df_filtrado = filtros(df_dados, key)
    st.caption(f"{len(df_filtrado)} registro(s) exibidos")

    df_edit = montar_tabela_editavel(df_filtrado, df_controle)

    colunas_fixas = [c for c in df_edit.columns if c not in ('Tratado', 'Observacao_Controle')]

    resultado = st.data_editor(
        df_edit,
        column_config={
            'Tratado': st.column_config.CheckboxColumn("✅ Tratado", default=False),
            'Observacao_Controle': st.column_config.TextColumn(
                "📝 Observação",
                help="Campo livre para anotações. Será salvo e compartilhado.",
                max_chars=500,
            ),
        },
        disabled=colunas_fixas,
        hide_index=True,
        use_container_width=True,
        key=f"editor_{key}",
    )

    col_salvar, col_export = st.columns([1, 4])
    with col_salvar:
        if st.button("💾 Salvar alterações", key=f"salvar_{key}", type="primary"):
            if resultado is not None and 'Chave de Acesso' in resultado.columns:
                agora = datetime.now().strftime("%d/%m/%Y %H:%M")

                # Monta registros novos a partir da edição
                novos = resultado[['Chave de Acesso', 'Tratado', 'Observacao_Controle']].copy()
                novos.columns = ['Chave', 'Tratado', 'Observacao_Controle']
                novos['Responsavel']   = responsavel
                novos['Atualizado_em'] = agora

                # Mescla com controle existente (preserva outras chaves)
                chaves_editadas = set(novos['Chave'].tolist())
                controle_restante = df_controle[~df_controle['Chave'].isin(chaves_editadas)]
                df_novo_controle = pd.concat([controle_restante, novos], ignore_index=True)

                salvar_controle(df_novo_controle)
                buscar_planilha_bytes.clear()
                st.rerun()

    with col_export:
        btn_exportar(df_filtrado, nome_export)


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
def main():
    # ── Sidebar ──────────────────────────────────────────
    with st.sidebar:
        st.header("Controles")

        responsavel = st.text_input(
            "👤 Seu nome",
            key="responsavel",
            placeholder="Ex: Ana Silva",
            help="Identifica quem fez cada alteração no controle compartilhado.",
        )
        if not responsavel:
            st.warning("Informe seu nome para salvar alterações.")

        st.divider()

        if st.button("🔄 Atualizar dados", type="primary", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        if 'atualizado_em' in st.session_state:
            st.caption(f"Atualizado em: {st.session_state['atualizado_em']}")

        st.divider()
        st.caption("**Filtro base (SAP)**")
        st.caption("Tipo Doc. Fiscal: NF-e")
        st.caption("Situação: Não informado / Uso autorizado / Doc não gerado")
        st.divider()
        st.caption("**Excluídos por palavras em:**")
        st.caption("Observação · Natureza Operação · Tipo Documento SAP")
        st.caption("*(comodato, retorno, devolução)*")

    # ── Carrega dados ─────────────────────────────────────
    with st.spinner("Buscando e cruzando dados…"):
        raw          = buscar_planilha_bytes()
        df_merge, df_pend, df_fora, df_excl = processar(raw)
        df_controle  = carregar_controle()
        st.session_state['atualizado_em'] = datetime.now().strftime("%d/%m/%Y %H:%M")

    # ── KPIs ─────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total no escopo",      len(df_merge))
    c2.metric("⚠️ Pendências",         len(df_pend))
    c3.metric("❌ Não gerados no SAP", int((df_merge.get('Status Cruzamento', pd.Series()) == 'Não gerado no SAP').sum()))
    c4.metric("📦 Excluídos",          len(df_excl))

    st.divider()

    # ── Abas ─────────────────────────────────────────────
    t1, t2, t3, t4 = st.tabs([
        f"⚠️ Pendências ({len(df_pend)})",
        f"📄 Cruzamento SAP × Qive ({len(df_merge)})",
        f"📂 Fora do Escopo ({len(df_fora)})",
        f"📦 Excluídos ({len(df_excl)})",
    ])

    with t1:
        st.info("Marque **Tratado** e adicione observações. Clique em **Salvar alterações** para compartilhar com a equipe.")
        render_editavel(df_pend, df_controle, "pend", "pendencias", responsavel)

    with t2:
        st.info("Visão completa do cruzamento. Marcações e observações também são salvas aqui.")
        render_editavel(df_merge, df_controle, "merge", "cruzamento_sap_qive", responsavel)

    with t3:
        st.caption("Documentos SAP fora dos critérios (tipo ou status não elegíveis).")
        st.dataframe(df_fora, hide_index=True, use_container_width=True)
        btn_exportar(df_fora, "fora_do_escopo")

    with t4:
        st.caption("Notas removidas por conter comodato, retorno ou devolução.")
        st.dataframe(df_excl, hide_index=True, use_container_width=True)
        btn_exportar(df_excl, "excluidos")


if __name__ == "__main__":
    main()