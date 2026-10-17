"""Dashboard Streamlit – Levantamento de BD (CDVs e Sinais, Vias 1 e 2).

Rodar:  streamlit run app.py
Coloque os dois .xlsx na mesma pasta deste arquivo (ou envie pela barra lateral).
"""
import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Levantamento BD • Vias 1 e 2", page_icon="🚆",
                   layout="wide", initial_sidebar_state="collapsed")
st.markdown("""<style>
.block-container{padding-top:1.5rem;max-width:1500px}
[data-testid=stMetricValue]{font-size:2.4rem;font-weight:700}
[data-testid=stMetricLabel]{font-size:1rem}
</style>""", unsafe_allow_html=True)

BASE = Path(__file__).parent
ARQ_CDV = "levantamento_BD_CDVs_via_1_e_2.xlsx"
ARQ_SIN = "levantamento_BD_Sinais_via_1_e_2.xlsx"
COR = {"Preenchido": "#2E9E6B", "Não utilizado": "#B8C0CC", "Pendente": "#E8833A"}
FUNC_CDV = [("Codificação", "NCTP"), ("Codificação", "SCTP"), ("Codificação", "N12-CTP"),
            ("Indicação", "TPR"), ("Indicação", "TPRN")]
FUNC_SIN = [("Vermelho", "RGE"), ("Vermelho", "RGENX"), ("Amarelo", "YGE"), ("Amarelo", "YGENX")]
PADRAO = {"Locação": re.compile(r"^L[BC] [A-Z]{1,2}-\d+$"),
          "Armário TB": re.compile(r"^TB\d-[A-Z]\d+$"),
          "Interface": re.compile(r"^B\d{2} \d{2}$")}
# Legenda de siglas. Edite os textos conforme o padrão da sua equipe.
GLOSSARIO = {
    "CDV": ("Equipamento", "Circuito de via: trecho de via monitorado para detectar a presença de trem."),
    "Sinal": ("Equipamento", "Sinal luminoso da via, com aspectos vermelho e amarelo."),
    "NCTP": ("Codificação", "Ponto de codificação do CDV (lado Norte)."),
    "SCTP": ("Codificação", "Ponto de codificação do CDV (lado Sul)."),
    "N12-CTP": ("Codificação", "Ponto de codificação do CDV (N12)."),
    "TPR": ("Indicação", "Ponto de indicação do CDV."),
    "TPRN": ("Indicação", "Ponto de indicação do CDV (variante N)."),
    "RGE": ("Aspecto vermelho", "Ponto do aspecto vermelho do sinal."),
    "RGENX": ("Aspecto vermelho", "Ponto do aspecto vermelho do sinal (variante NX)."),
    "YGE": ("Aspecto amarelo", "Ponto do aspecto amarelo do sinal."),
    "YGENX": ("Aspecto amarelo", "Ponto do aspecto amarelo do sinal (variante NX)."),
    "Locação": ("Local", "Posição onde o ponto é ligado, ex.: LB A-17 ou LC D-03."),
    "Armário TB": ("Local", "Armário de bornes, ex.: TB2-E5."),
    "Interface": ("Local", "Ponto de interface, ex.: B03 21."),
    "Bastidor": ("Local", "Bastidor, ex.: DCC1."),
    "Preenchido": ("Situação", "Ponto informado na planilha."),
    "Não utilizado": ("Situação", "Célula com “-”: o ponto não é usado."),
    "Pendente": ("Situação", "Célula vazia: falta informar o ponto."),
}
CAT = {fun: cat for cat, fun in FUNC_CDV + FUNC_SIN}


# ───────────────────────── leitura e tratamento ─────────────────────────
def txt(v):
    return "" if pd.isna(v) else str(v).strip()


def status_de(v):  # vazio = pendente | "-" = não utilizado
    return "Pendente" if v == "" else "Não utilizado" if v == "-" else "Preenchido"


def tipo_local(loc):
    l = loc.lower()
    for ini, t in (("locação", "Locação"), ("armári", "Armário TB"), ("armari", "Armário TB"),
                   ("interface", "Interface"), ("bastidor", "Bastidor")):
        if l.startswith(ini):
            return t
    return "Outro"


def fora_padrao(tipo, ponto):
    p = PADRAO.get(tipo)
    if not p or ponto in ("", "-"):
        return False
    return not all(p.match(x.strip()) for x in ponto.split("/"))


def familia(tipo, ponto):
    if ponto in ("", "-"):
        return ""
    p = ponto.split("/")[0].strip()
    if tipo == "Locação":
        return p[:2]
    if tipo == "Armário TB":
        m = re.match(r"TB\d*", p)
        return m.group() if m else "Fora do padrão"
    return tipo


def parse(raw, ativo, prefixo, cols, funcs, dados):
    """Lê os blocos lado a lado e devolve uma tabela 'longa' (1 linha por ponto)."""
    regs, cont = [], {}
    for c0 in cols:
        r = 0
        while r < len(raw):
            cel = txt(raw.iat[r, c0])
            if not cel.upper().startswith(prefixo):
                r += 1
                continue
            nome = re.sub(r"\s*LUZ$", "", cel[len(prefixo):].strip(), flags=re.I).strip()
            cont[nome] = cont.get(nome, 0) + 1
            bid = nome if cont[nome] == 1 else f"{nome} (#{cont[nome]})"
            m = re.match(r"^(\d)([NS])", nome) if ativo == "CDV" else None
            via = f"Via {m[1]}" if m else "—"
            sent = {"N": "Norte", "S": "Sul"}[m[2]] if m else "—"
            codigos = [txt(raw.iat[r + 3, c0 + 1 + k]) for k in range(len(funcs))]
            rr = r + dados
            while rr < len(raw):
                loc = txt(raw.iat[rr, c0])
                if not loc or loc.upper().startswith(prefixo):
                    break
                tipo = tipo_local(loc)
                local = loc if tipo in ("Locação", "Bastidor") else tipo
                for k, (cat, fun) in enumerate(funcs):
                    ponto = txt(raw.iat[rr, c0 + 1 + k])
                    cod = codigos[k] or fun
                    if ativo == "Sinal" and cod == fun:
                        cod = nome + fun
                    regs.append(dict(ativo=ativo, id=bid, nome=nome, via=via, sentido=sent,
                                     local=local, tipo_local=tipo, categoria=cat, funcao=fun,
                                     codigo=cod, ponto=ponto, status=status_de(ponto),
                                     familia=familia(tipo, ponto),
                                     fora_padrao=fora_padrao(tipo, ponto)))
                rr += 1
            r = rr
    for x in regs:
        x["duplicado"] = cont[x["nome"]] > 1
    return pd.DataFrame(regs)


@st.cache_data
def ler(b):
    return pd.read_excel(BytesIO(b), header=None)


def carregar(arq, rotulo):
    p = BASE / arq
    if p.exists():
        return p.read_bytes()
    up = st.sidebar.file_uploader(rotulo, type="xlsx")
    return up.getvalue() if up else None


b_cdv, b_sin = carregar(ARQ_CDV, "Planilha de CDVs"), carregar(ARQ_SIN, "Planilha de Sinais")
if b_cdv is None or b_sin is None:
    st.info("Coloque os dois arquivos .xlsx na mesma pasta do app ou envie-os pela barra lateral.")
    st.stop()

df = pd.concat([parse(ler(b_cdv), "CDV", "CDV ", (1, 8), FUNC_CDV, 4),
                parse(ler(b_sin), "Sinal", "SINAL ", (0, 6, 12, 18), FUNC_SIN, 5)],
               ignore_index=True)

# ───────────────────────── filtros ─────────────────────────
with st.sidebar:
    st.header("Filtros")
    vias_all = sorted(df.via[df.ativo == "CDV"].unique())
    vias = st.multiselect("Via (CDVs)", vias_all, default=vias_all)
    locs_all = sorted(df.tipo_local.unique())
    locs = st.multiselect("Tipo de local", locs_all, default=locs_all)
f = df[df.tipo_local.isin(locs) & ((df.ativo != "CDV") | df.via.isin(vias))]
cdv, sin = f[f.ativo == "CDV"], f[f.ativo == "Sinal"]

# ───────────────────────── cabeçalho / KPIs ─────────────────────────
st.title("🚆 Levantamento de BD — Vias 1 e 2")
st.caption("CDVs e Sinais · **Preenchido** = ponto informado · **Não utilizado** = “-” · "
           "**Pendente** = célula vazia")

pre, pen = (f.status == "Preenchido").sum(), (f.status == "Pendente").sum()
taxa = pre / (pre + pen) if (pre + pen) else 0
k = st.columns(5)
k[0].metric("CDVs mapeados", cdv.nome.nunique())
k[1].metric("Sinais mapeados", sin.nome.nunique())
k[2].metric("Pontos preenchidos", f"{pre:,}".replace(",", "."))
k[3].metric("Taxa de preenchimento", f"{taxa:.1%}")
k[4].metric("Pendências", int(pen))


def mostrar(fig, h=360):
    fig.update_layout(template="plotly_white", height=h, font_size=14, legend_title_text="",
                      margin=dict(t=20, b=10, l=10, r=10))
    st.plotly_chart(fig)


def legenda(chaves, aberta=False):
    with st.expander("📖 Legenda de siglas", expanded=aberta):
        st.table(pd.DataFrame([(k, *GLOSSARIO[k]) for k in chaves],
                              columns=["Sigla", "Grupo", "Significado"]).set_index("Sigla"))
        st.caption("Significados a confirmar com a equipe (editáveis no dicionário GLOSSARIO do código).")


def aba_ativo(d, nome, ordem, siglas):
    if d.empty:
        st.info("Sem dados para os filtros selecionados.")
        return
    ok = d[d.status == "Preenchido"]
    st.markdown(f"#### 1. Quantos pontos foram preenchidos em cada função?")
    st.caption("Cada barra é uma função do " + nome + ". Quanto maior a barra, mais pontos informados.")
    q = ok.groupby("funcao").size().reindex(ordem, fill_value=0).reset_index(name="pontos")
    q["rotulo"] = q.funcao.map(lambda x: f"{x}<br>({CAT[x]})")
    fig = px.bar(q, x="rotulo", y="pontos", text="pontos", color_discrete_sequence=["#2a78d6"])
    fig.update_xaxes(title=None, categoryorder="array", categoryarray=list(q.rotulo))
    fig.update_yaxes(title=None)
    mostrar(fig, 320)
    st.markdown(f"#### 2. Como está cada {nome}?")
    st.caption("Cada linha é um " + nome + " e cada número é a quantidade de pontos preenchidos naquela função. "
               "Cor mais escura = mais pontos. **0** = nenhum ponto preenchido (ou função não usada).")
    h = ok.pivot_table(index="id", columns="funcao", values="ponto", aggfunc="count",
                       fill_value=0).reindex(columns=ordem, fill_value=0).sort_index()
    if not h.empty:
        fig = px.imshow(h, text_auto=True, aspect="auto", color_continuous_scale="Blues")
        fig.update_layout(coloraxis_showscale=False)
        fig.update_xaxes(side="top", title=None)
        fig.update_yaxes(title=None)
        mostrar(fig, max(320, 26 * len(h) + 80))
    legenda(siglas)


# ───────────────────────── abas ─────────────────────────
t1, t2, t3, t4, t5 = st.tabs(["Visão geral", "CDVs", "Sinais", "Qualidade dos dados", "Dados"])

with t1:
    st.markdown("#### 1. Os pontos estão preenchidos?")
    st.caption(f"**{taxa:.0%} dos pontos que deveriam ser informados já estão preenchidos.** "
               "Pontos com “-” não entram na conta, pois não são usados.")
    a = f.groupby(["ativo", "status"]).size().reset_index(name="pontos")
    a["pct"] = a.pontos / a.groupby("ativo").pontos.transform("sum")
    fig = px.bar(a, y="ativo", x="pct", color="status", orientation="h", color_discrete_map=COR,
                 text=a.pct.map(lambda v: f"{v:.0%}" if v >= 0.03 else ""),
                 category_orders={"status": list(COR)})
    fig.update_xaxes(tickformat=".0%", title=None)
    fig.update_yaxes(title=None)
    mostrar(fig, 240)
    st.markdown("#### 2. Onde os pontos estão ligados?")
    st.caption("Quantidade de pontos preenchidos em cada tipo de local, separada por CDV e Sinal.")
    b = f[f.status == "Preenchido"].groupby(["tipo_local", "ativo"]).size().reset_index(name="pontos")
    fig = px.bar(b, y="tipo_local", x="pontos", color="ativo", barmode="group", text="pontos",
                 color_discrete_sequence=["#2a78d6", "#eb6834"])
    fig.update_xaxes(title=None)
    fig.update_yaxes(title=None, categoryorder="total ascending")
    mostrar(fig, 320)
    legenda(["CDV", "Sinal", "Locação", "Armário TB", "Interface", "Bastidor",
             "Preenchido", "Não utilizado", "Pendente"], aberta=True)

with t2:
    aba_ativo(cdv, "CDV", [x[1] for x in FUNC_CDV],
              ["CDV", "NCTP", "SCTP", "N12-CTP", "TPR", "TPRN", "Locação", "Armário TB", "Interface"])
with t3:
    aba_ativo(sin, "Sinal", [x[1] for x in FUNC_SIN],
              ["Sinal", "RGE", "RGENX", "YGE", "YGENX", "Locação", "Armário TB", "Interface", "Bastidor"])

with t4:
    dup = f[f.duplicado].groupby(["ativo", "nome"]).id.nunique().reset_index(name="ocorrências")
    q = st.columns(3)
    q[0].metric("Células pendentes", int(pen))
    q[1].metric("Pontos fora do padrão", int(f.fora_padrao.sum()))
    q[2].metric("Blocos repetidos", len(dup))
    st.markdown("##### Blocos com o mesmo nome aparecendo mais de uma vez")
    st.dataframe(dup.rename(columns={"ativo": "Ativo", "nome": "Nome"}), hide_index=True)
    cols = ["ativo", "id", "local", "funcao", "codigo", "ponto"]
    st.markdown("##### Pontos fora do padrão de escrita (revisar)")
    st.dataframe(f[f.fora_padrao][cols], hide_index=True)
    st.markdown("##### Pendências (células vazias)")
    st.dataframe(f[f.status == "Pendente"][cols[:-1]], hide_index=True)

with t5:
    st.dataframe(f.drop(columns=["nome", "duplicado"]), hide_index=True)
    st.download_button("⬇️ Baixar dados tratados (CSV)", f.to_csv(index=False).encode("utf-8-sig"),
                       "levantamento_bd_tratado.csv", "text/csv")
