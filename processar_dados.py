"""
processar_dados.py v5 â Dashboard AnalÃ­tica de CX Â· Xtay
MÃ³dulo reutilizÃ¡vel: processar(rows_droz, rows_occ) â payload dict
Uso local: python3 processar_dados.py  (lÃª Base de Atendimento.xlsx)
"""
import json, re
from collections import defaultdict, Counter
from datetime import datetime, timedelta

ARQUIVO_ENTRADA  = "Base de Atendimento.xlsx"
ARQUIVO_TEMPLATE = "template.html"
ARQUIVO_SAIDA    = "dashboard_cx.html"

TAGS_STATUS_CLIENTE = {
    "esta_hospedado","nao_esta_hospedado",
    "nao_possui_checkin_hoje_nem_futuros","existe_checkin_futuro",
}
MAPA_PRIORIDADE = {"prioridade_alta":"Alta","prioridade_media":"MÃ©dia","prioridade_baixa":"Baixa"}

EMPREEND_KEYWORDS = [
    ("linked batel","Linked Batel"),("linked_batel","Linked Batel"),
    ("campos salles","Campos Sales"),("campos_salles","Campos Sales"),
    ("campos sales","Campos Sales"),("campos_sales","Campos Sales"),
    ("upper west","Upper West"),("upper_west","Upper West"),
    ("simple smart","Simple Smart"),("simple_smart","Simple Smart"),
    ("vila madalena","Vila Madalena"),("vila_madalena","Vila Madalena"),
    ("the bridge","The Bridge"),("the_bridge","The Bridge"),
    ("ipiranga","Ipiranga"),
    ("upper","Upper West"),("simple","Simple Smart"),
    ("cais","Cais"),("atrium","Atrium"),("fuji","Fuji"),
    ("soho","Soho"),("oslo","Oslo"),("puc","PUC"),
]

CLUSTERS = {
    "Vila Madalena":"SÃ£o Paulo","Upper West":"SÃ£o Paulo","Ipiranga":"SÃ£o Paulo",
    "Cais":"Porto Alegre","Simple Smart":"JoÃ£o Pessoa",
    "Campos Sales":"Curitiba","PUC":"Curitiba",
    "Linked Batel":"Linked Batel",
    "Atrium":"Santa Catarina","Fuji":"Santa Catarina",
    "Soho":"Santa Catarina","Oslo":"Santa Catarina","The Bridge":"Santa Catarina",
}

CLASSIF_EXACT = {
    "comercial":"fila_comercial","atendimento-comercial":"fila_comercial",
    "hospedar":"fila_comercial","morar":"fila_comercial",
    "comercial_hospedar":"fila_comercial","comercial_morar":"fila_comercial",
    "checkin_b2b":"fila_comercial","cupom_de_desconto_reservas":"fila_comercial",
    "reserva_upgrade":"fila_comercial","novos_negocios":"fila_comercial",
    "novos_negocios_marketing":"fila_comercial",
    "duvidas_gerais":"duvidas","duvidas_predio":"duvidas",
    "check_in_duvidas":"duvidas","early_check_in_duvidas":"duvidas",
    "late_check_out_duvidas":"duvidas","duvidas_gerais_senha_de_acesso":"duvidas",
    "detalhes_da_acomodacao":"duvidas","informacoes_de_acesso_wi_fi":"duvidas",
    "estacionamento_duvidas":"duvidas","estacionamento_reserva":"duvidas",
    "estacionamento_dificuldades_de_acesso":"duvidas","cafe_da_manha":"duvidas",
    "maleiro":"duvidas","mini_mercado":"duvidas","lavanderia_omo":"duvidas",
    "duvidas_sobre_limpeza":"duvidas","checkout_duvidas":"duvidas",
    "confirmacao_de_reserva":"duvidas","duvidas_cadastro_visitante":"duvidas",
    "cadastro_visitante":"duvidas","feedback_estadia":"duvidas",
    "suspeita_de_fraude":"duvidas","alteracao_de_datas":"duvidas",
    "reembolso":"duvidas","solicitacao_de_nota_fiscal":"duvidas",
    "cancelamento_de_reserva_fora_da_politica":"duvidas",
    "cancelamento_de_reserva_dentro_da_politica":"duvidas",
    "alteracao_de_reserva_prorrogacao":"duvidas",
    "alteracao_de_reserva_inclusao_de_dependente":"duvidas",
    "alteracao_de_reserva_inclusao_de_acompanhante":"duvidas",
    "solicitacao_de_early_check_in_cortesia":"duvidas",
    "solicitacao_de_early_check_in_autorizado_com_custos":"duvidas",
    "solicitacao_de_early_check_in_sem_disponibilidade":"duvidas",
    "duvidas_omnibees":"duvidas","reclamacao_da_vista":"duvidas",
    "duvidas_descarte_lixo":"duvidas","check_in_de_dependente":"duvidas",
    "inclusao_check_in_de_dependente":"duvidas","alteracao_de_dados_cadastrais":"duvidas",
    "problemas_com_senha_da_porta":"tech","problemas_com_a_facial":"tech",
    "check_in_erro_de_verificacao":"tech","check_in_manual_dificuldades_do_hospede":"tech",
    "check_in_manual_documento_digital":"tech","check_in_reserva_nao_localizada":"tech",
    "abertura_remota_check_out_vencido":"tech","abertura_remota_dificuldades_do_hospede":"tech",
    "problemas_com_link_de_pagamento":"tech","problemas_com_pagamento":"tech","checkout_manual":"tech",
    "ar_condicionado":"operacional","hidraulica":"operacional",
    "aviso_de_falta_de_energia":"operacional","reclamacao_de_internet":"operacional",
    "troca_de_apartamento":"operacional","achados_e_perdidos":"operacional",
    "reposicao_de_amenidades":"operacional","item_faltante_na_unidade":"operacional",
    "item_danificado_no_apto":"operacional","reclamacao_de_limpeza":"operacional",
    "reclamacao_de_barulho":"operacional","reclamacao_estadia":"operacional",
    "limpeza_nao_realizada":"operacional",
    "solicitacao_de_late_check_out_cortesia":"operacional",
    "solicitacao_de_late_check_out_autorizado_com_custos":"operacional",
    "solicitacao_de_late_check_out_sem_disponibilidade":"operacional",
    "mkt_hospede_frequent":"marketing","hospede_mkt":"marketing",
    "sem_retorno":"improdutivos","falta_de_interacao":"improdutivos",
    "teste":"improdutivos","trabalhe_conosco":"improdutivos",
    "contato_ativo":"contato_ativo","aviso_de_check_out":"contato_ativo",
    "conversa_operacional":"contato_ativo",
}

CLASSIF_PREFIXOS = [
    ("mensalista_","fila_comercial"),("cotacao_orcamento_","fila_comercial"),
    ("novos_negocios","fila_comercial"),("solicitacao_de_servico_","operacional"),
    ("abertura_remota_","tech"),("check_in_manual_","tech"),
    ("cancelamento_de_reserva_","duvidas"),("alteracao_de_reserva_","duvidas"),
    ("solicitacao_de_early_check_in_","duvidas"),("solicitacao_de_late_check_out_","operacional"),
]

CLASSIF_META = {
    "fila_comercial":{"label":"Fila comercial","cor":"#FF6A13"},
    "duvidas":{"label":"DÃºvidas e solicitaÃ§Ãµes gerais","cor":"#4F79E4"},
    "tech":{"label":"Tech","cor":"#9B59B6"},
    "operacional":{"label":"SolicitaÃ§Ãµes operacionais","cor":"#E67E22"},
    "marketing":{"label":"Marketing","cor":"#27AE60"},
    "improdutivos":{"label":"Improdutivos","cor":"#95A5A6"},
    "contato_ativo":{"label":"Contato ativo","cor":"#1ABC9C"},
}
CLASSIF_ORDER = ["fila_comercial","duvidas","tech","operacional","marketing","improdutivos","contato_ativo"]

EMPREEND_ORDER = [
    "Vila Madalena","Upper West","Ipiranga",
    "Cais","PUC","Campos Sales","Linked Batel",
    "Atrium","Fuji","Soho","Oslo","The Bridge","Simple Smart",
]
CLUSTER_ORDER = ["SÃ£o Paulo","Curitiba","Linked Batel","Porto Alegre","Santa Catarina","JoÃ£o Pessoa"]

NM = {"01":"Jan","02":"Fev","03":"Mar","04":"Abr","05":"Mai","06":"Jun",
      "07":"Jul","08":"Ago","09":"Set","10":"Out","11":"Nov","12":"Dez"}

MES_NUM = {
    "jan":"01","fev":"02","mar":"03","abr":"04","mai":"05","jun":"06",
    "jul":"07","ago":"08","set":"09","out":"10","nov":"11","dez":"12",
    "janeiro":"01","fevereiro":"02","marco":"03","abril":"04","maio":"05",
    "junho":"06","julho":"07","agosto":"08","setembro":"09","outubro":"10",
    "novembro":"11","dezembro":"12",
}

LABELS = {
    "check_in_duvidas":"Check-in â dÃºvidas","duvidas_gerais":"DÃºvidas gerais",
    "sem_retorno":"Sem retorno","check_in_manual_dificuldades_do_hospede":"Check-in manual â dificuldades",
    "alteracao_de_reserva_prorrogacao":"ProrrogaÃ§Ã£o de reserva",
    "problemas_com_senha_da_porta":"Problemas com senha da porta",
    "solicitacao_de_early_check_in_cortesia":"Early check-in (cortesia)",
    "solicitacao_de_servico_manutencao":"ManutenÃ§Ã£o",
    "early_check_in_duvidas":"Early check-in â dÃºvidas",
    "estacionamento_duvidas":"Estacionamento â dÃºvidas",
    "solicitacao_de_late_check_out_cortesia":"Late check-out (cortesia)",
    "troca_de_apartamento":"Troca de apartamento",
    "cotacao_orcamento_venda_nao_finalizada":"CotaÃ§Ã£o nÃ£o finalizada",
    "cotacao_orcamento_venda_finalizada":"CotaÃ§Ã£o finalizada",
    "inclusao_check_in_de_dependente":"InclusÃ£o dependente no check-in",
    "duvidas_gerais_senha_de_acesso":"DÃºvidas senha de acesso",
    "problemas_com_a_facial":"Problemas com facial",
    "mensalista_agendamento_de_limpeza_quinzenal":"Mensalista â limpeza quinzenal",
    "achados_e_perdidos":"Achados e perdidos","solicitacao_de_nota_fiscal":"Nota fiscal",
    "solicitacao_de_servico_item_extra":"Item extra",
    "check_in_reserva_nao_localizada":"Check-in â reserva nÃ£o localizada",
    "confirmacao_de_reserva":"ConfirmaÃ§Ã£o de reserva",
    "solicitacao_de_servico_limpeza":"SolicitaÃ§Ã£o de limpeza",
    "reclamacao_de_limpeza":"ReclamaÃ§Ã£o de limpeza",
    "late_check_out_duvidas":"Late check-out â dÃºvidas",
    "mensalista_cotacao_e_duvidas":"Mensalista â cotaÃ§Ã£o",
    "check_in_de_dependente":"Check-in de dependente",
    "cafe_da_manha":"CafÃ© da manhÃ£","duvidas_sobre_limpeza":"DÃºvidas sobre limpeza",
    "alteracao_de_datas":"AlteraÃ§Ã£o de datas",
    "cancelamento_de_reserva_fora_da_politica":"Cancelamento fora da polÃ­tica",
    "cancelamento_de_reserva_dentro_da_politica":"Cancelamento dentro da polÃ­tica",
    "solicitacao_de_servico_troca_de_toalhas":"Troca de toalhas",
    "solicitacao_de_early_check_in_autorizado_com_custos":"Early check-in (com custo)",
    "solicitacao_de_late_check_out_autorizado_com_custos":"Late check-out (com custo)",
    "problemas_com_pagamento":"Problemas com pagamento",
    "detalhes_da_acomodacao":"DÃºvidas sobre acomodaÃ§Ã£o",
    "problemas_com_link_de_pagamento":"Problemas com link de pagamento",
    "reembolso":"Reembolso","estacionamento_dificuldades_de_acesso":"Estacionamento â acesso",
    "reclamacao_de_internet":"ReclamaÃ§Ã£o de internet","item_danificado_no_apto":"Item danificado",
    "reclamacao_estadia":"ReclamaÃ§Ã£o de estadia","comercial":"Comercial",
    "atendimento-comercial":"Atendimento comercial","hospedar":"Hospedar","morar":"Morar",
    "contato_ativo":"Contato ativo","falta_de_interacao":"Falta de interaÃ§Ã£o",
    "aviso_de_check_out":"Aviso de check-out","checkout_manual":"Checkout manual",
    "check_in_erro_de_verificacao":"Check-in â erro de verificaÃ§Ã£o",
    "check_in_manual_documento_digital":"Check-in manual â doc digital",
    "suspeita_de_fraude":"Suspeita de fraude",
    "abertura_remota_check_out_vencido":"Abertura remota (check-out vencido)",
    "abertura_remota_dificuldades_do_hospede":"Abertura remota â dificuldades",
    "solicitacao_de_early_check_in_sem_disponibilidade":"Early check-in (sem disponibilidade)",
    "estacionamento_reserva":"Estacionamento â reserva","maleiro":"Maleiro",
    "mensalista_solicitacao_de_link_de_pagamento":"Mensalista â link de pagamento",
    "item_faltante_na_unidade":"Item faltante",
    "solicitacao_de_servico_troca_de_enxoval":"Troca de enxoval",
    "reposicao_de_amenidades":"ReposiÃ§Ã£o de amenidades",
    "reclamacao_de_barulho":"ReclamaÃ§Ã£o de barulho","limpeza_nao_realizada":"Limpeza nÃ£o realizada",
    "ar_condicionado":"Ar condicionado","hidraulica":"HidrÃ¡ulica",
    "aviso_de_falta_de_energia":"Falta de energia","feedback_estadia":"Feedback de estadia",
    "mkt_hospede_frequent":"HÃ³spede frequente","novos_negocios":"Novos negÃ³cios",
    "trabalhe_conosco":"Trabalhe conosco","conversa_operacional":"Conversa operacional",
    "teste":"Teste","checkin_b2b":"Check-in B2B","reserva_upgrade":"Upgrade de reserva",
    "cupom_de_desconto_reservas":"Cupom de desconto","mensalista_contrato_venda":"Mensalista â contrato",
    "solicitacao_de_late_check_out_sem_disponibilidade":"Late check-out (sem disponibilidade)",
    "duvidas_omnibees":"DÃºvidas Omnibees","alteracao_de_dados_cadastrais":"AlteraÃ§Ã£o de dados cadastrais",
}

# ââ FunÃ§Ãµes auxiliares âââââââââââââââââââââââââââââââââââââââ

def lbl(tag):
    t = tag.lower().strip()
    return LABELS.get(t, t.replace("_"," ").replace("-"," ").title())

def _v(val):
    if val is None: return None
    s = str(val).strip()
    return None if s == '' else s

def parse_min(val):
    v = _v(val)
    if not v: return None
    try:
        p = v.split(":")
        if len(p)==3: return int(p[0])*60+int(p[1])+float(p[2])/60
    except: pass
    return None

def parse_dt(val):
    v = _v(val)
    if not v: return None
    for f in ("%d/%m/%Y","%Y-%m-%d"):
        try: return datetime.strptime(v[:10],f).date()
        except: pass
    return None

def sem_lbl(d):
    return (d-timedelta(days=d.weekday())).strftime("%d/%m")

def avg(lst):
    lst = [x for x in lst if x is not None and x>=0]
    return round(sum(lst)/len(lst),1) if lst else None

def detect_emp(tag):
    t = tag.lower()
    for kw, nome in EMPREEND_KEYWORDS:
        if kw in t: return nome
    return None

def classify(tag):
    t = tag.lower().strip()
    cids = set()
    for p, cid in CLASSIF_PREFIXOS:
        if t.startswith(p): cids.add(cid)
    if t in CLASSIF_EXACT: cids.add(CLASSIF_EXACT[t])
    return cids

def _safe_float(val):
    v = _v(val)
    if not v: return 0.0
    # Remove %, R$, espaÃ§os nÃ£o-separÃ¡veis antes de converter
    cleaned = str(v).replace(',','.').replace('%','').replace('R$','').replace('\xa0','').strip()
    try: return float(cleaned)
    except: return 0.0

def _safe_int(val):
    v = _v(val)
    if not v: return 0
    cleaned = str(v).replace(',','.').replace('%','').strip()
    try: return int(float(cleaned))
    except: return 0

def _parse_mes_key(val, ano_base="2026"):
    """Normaliza label de mÃªs para YYYY-MM. Ex: 'Abril' â '2026-04'"""
    s = str(val or "").strip().lower()
    if not s: return None
    # JÃ¡ no formato YYYY-MM
    m = re.match(r'(\d{4})-(\d{2})', s)
    if m: return f"{m.group(1)}-{m.group(2)}"
    # Formato MM/YYYY ou MM-YYYY
    m = re.match(r'(\d{1,2})[/-](\d{4})', s)
    if m: return f"{m.group(2)}-{m.group(1).zfill(2)}"
    # Nome do mÃªs (com ou sem ano)
    for nome, num in MES_NUM.items():
        if s.startswith(nome):
            yr_m = re.search(r'\d{2,4}', s[len(nome):])
            yr = yr_m.group() if yr_m else ano_base
            if len(yr) == 2: yr = "20" + yr
            return f"{yr}-{num}"
    return None

def _build_metrics(tids, ticket_base, ticket_tags, occ_lookup=None):
    """Agrega mÃ©tricas para um conjunto de ticket IDs."""
    cc = defaultdict(int)
    ctpr = defaultdict(list); cttf = defaultdict(list); ctags = defaultdict(Counter)
    emp_c = Counter(); emp_m = defaultdict(Counter); clu_c = Counter()
    ag_c = defaultdict(int); ag_tpr = defaultdict(list)
    ag_ttf = defaultdict(list); ag_msg = defaultdict(list)
    pri = Counter(); tprs = []; ttfs = []

    for tid in tids:
        base = ticket_base[tid]
        tags = [t.strip() for t in ticket_tags.get(tid, [])]
        tpr = base["tpr"]; ttf = base["ttf"]
        if tpr is not None: tprs.append(tpr)
        if ttf is not None: ttfs.append(ttf)
        ag = base["agente"] or "Sem agente"
        ag_c[ag]+=1; ag_tpr[ag].append(tpr); ag_ttf[ag].append(ttf); ag_msg[ag].append(base["msgs"])
        emps = set()
        for tag in tags:
            e = detect_emp(tag)
            if e: emps.add(e)
        for e in emps:
            emp_c[e]+=1; clu_c[CLUSTERS.get(e,"Outros")]+=1
        for tag in tags:
            t = tag.lower().strip()
            if t in MAPA_PRIORIDADE: pri[MAPA_PRIORIDADE[t]]+=1; break
        mot = [t for t in tags if t.lower() not in TAGS_STATUS_CLIENTE
               and t.lower() not in MAPA_PRIORIDADE and not detect_emp(t)]
        for e in emps:
            for mt in mot: emp_m[e][lbl(mt)]+=1
        tclassifs = set()
        for tag in mot:
            for cid in classify(tag):
                tclassifs.add(cid); ctags[cid][lbl(tag)]+=1
        for cid in tclassifs:
            cc[cid]+=1; ctpr[cid].append(tpr); cttf[cid].append(ttf)

    classifs = [{"id":cid,"label":CLASSIF_META[cid]["label"],"cor":CLASSIF_META[cid]["cor"],
        "count":cc.get(cid,0),"tpr_med":avg(ctpr.get(cid,[])),"ttf_med":avg(cttf.get(cid,[])),
        "top_tags":[{"label":k,"count":v} for k,v in ctags[cid].most_common(8)]}
        for cid in CLASSIF_ORDER]

    emps_lista = []
    for nome in EMPREEND_ORDER:
        tickets = emp_c.get(nome,0)
        occ = (occ_lookup or {}).get(nome,{})
        reservas = occ.get("reservas",0)
        occ_pct = occ.get("occ_pct",None)
        taxa = round(tickets/reservas*100,1) if reservas>0 else None
        top8 = [{"label":k,"count":v} for k,v in emp_m[nome].most_common(8)]
        emps_lista.append({"nome":nome,"cluster":CLUSTERS.get(nome,"Outros"),
            "count":tickets,"top_motivos":top8,
            "reservas":reservas,"occ_pct":occ_pct,"taxa_contato":taxa})
    emps_lista.sort(key=lambda x: -x["count"])

    agentes = [{"nome":ag,"tickets":cnt,"tpr_med":avg(ag_tpr[ag]),
        "ttf_med":avg(ag_ttf[ag]),"msgs_med":avg(ag_msg[ag])}
        for ag,cnt in sorted(ag_c.items(),key=lambda x:-x[1])]

    clusters_lista = [{"nome":cl,"count":clu_c.get(cl,0)} for cl in CLUSTER_ORDER]

    return {
        "total": len(tids),
        "tpr_med": avg(tprs),
        "ttf_med": avg(ttfs),
        "classifs": classifs,
        "empreendimentos": emps_lista,
        "agentes": agentes,
        "clusters": clusters_lista,
        "prioridades": {"Alta":pri.get("Alta",0),"Media":pri.get("MÃ©dia",0),"Baixa":pri.get("Baixa",0)},
    }

# ââ FunÃ§Ã£o principal âââââââââââââââââââââââââââââââââââââââââ

def processar(rows_droz, rows_occ=None):
    """
    Processa dados de atendimento e OCC.
    rows_droz : lista de linhas com dados de tickets (sem header)
    rows_occ  : lista de linhas com dados de OCC (sem header), ou None
    Retorna: payload dict pronto para JSON
    """
    # 1. Leitura de tickets
    ticket_base = {}
    ticket_tags = defaultdict(list)
    for row in rows_droz:
        row = list(row) + [None]*25
        tid = _v(row[1])
        if not tid: continue
        tag = _v(row[23])
        if tid not in ticket_base:
            ticket_base[tid] = {
                "agente": _v(row[19]),
                "data":   parse_dt(row[22]),
                "tpr":    parse_min(row[12]),
                "ttf":    parse_min(row[13]),
                "msgs":   _safe_int(row[14]),
            }
        if tag: ticket_tags[tid].append(tag.strip())

    # 2. Volume por perÃ­odo (para construÃ§Ã£o de keys)
    vd = Counter(); vs = Counter(); vm = Counter()
    for tid, base in ticket_base.items():
        dt = base["data"]
        if dt:
            vd[dt.strftime("%Y-%m-%d")]+=1
            vm[dt.strftime("%Y-%m")]+=1
            vs[sem_lbl(dt)]+=1

    # 3. Leitura de OCC
    # Estrutura da aba "Reservas e OCC" (com coluna MÃªs em A):
    #   col A (0): mÃªs, col B (1): empreendimento
    #   col E (4): un. ocupadas, col G (6): faturamento, col H (7): occ%
    occ_raw_data = defaultdict(dict)
    if rows_occ:
        for row in rows_occ:
            row = list(row) + [None]*9
            if not _v(row[1]): continue
            mes_raw = str(_v(row[0]) or "").strip()
            canon = detect_emp(str(row[1]))
            if not canon: continue
            un_ocup = _safe_int(row[4])
            occ_r = _safe_float(row[7])
            occ_pct = round(occ_r*100, 1) if 0 < occ_r <= 1.0 else round(occ_r, 1)
            fat = _safe_float(row[6])
            occ_raw_data[mes_raw][canon] = {"reservas":un_ocup,"occ_pct":occ_pct,"faturamento":fat}

    # 4. Normaliza chaves de OCC para YYYY-MM
    ano_base = sorted(vm.keys())[0][:4] if vm else "2026"
    occ_data = defaultdict(dict)
    for mes_raw, mes_data in occ_raw_data.items():
        key = _parse_mes_key(mes_raw, ano_base) or mes_raw
        occ_data[key] = mes_data

    # 5. OCC agregado (todos os meses)
    occ_total = defaultdict(lambda: {"reservas":0,"occ_pct_vals":[],"faturamento":0.0})
    for mes_data in occ_data.values():
        for emp, vals in mes_data.items():
            occ_total[emp]["reservas"] += vals["reservas"]
            occ_total[emp]["faturamento"] += vals["faturamento"]
            if vals["occ_pct"]: occ_total[emp]["occ_pct_vals"].append(vals["occ_pct"])
    occ_agg = {}
    for emp, v in occ_total.items():
        pcts = v["occ_pct_vals"]
        occ_agg[emp] = {
            "reservas": v["reservas"],
            "occ_pct": round(sum(pcts)/len(pcts),1) if pcts else None,
            "faturamento": round(v["faturamento"],2),
        }

    # 6. MÃ©tricas globais
    all_tids = list(ticket_base.keys())
    gm = _build_metrics(all_tids, ticket_base, ticket_tags, occ_agg)

    # 7. Volume charts (global)
    dias_s = sorted(vd.items())
    seen_w = []; seen_ws = set()
    for d_str, _ in dias_s:
        d = datetime.strptime(d_str,"%Y-%m-%d").date()
        lw = sem_lbl(d)
        if lw not in seen_ws: seen_ws.add(lw); seen_w.append(lw)
    sems = [{"d":w,"c":vs[w]} for w in seen_w]
    meses_vol = [{"d":NM.get(m[5:7],m[5:7]),"c":c,"k":m} for m,c in sorted(vm.items())]

    # 8. MÃ©tricas por mÃªs
    por_mes = {}
    for mes_key in sorted(vm.keys()):
        month_tids = [tid for tid, base in ticket_base.items()
                      if base["data"] and base["data"].strftime("%Y-%m") == mes_key]
        month_occ = dict(occ_data.get(mes_key, {}))
        mm = _build_metrics(month_tids, ticket_base, ticket_tags, month_occ)
        month_dias = [{"d":d[8:]+"/"+d[5:7],"c":c}
                      for d,c in sorted(vd.items()) if d[:7]==mes_key]
        por_mes[mes_key] = {**mm, "label":NM.get(mes_key[5:7],mes_key), "dias":month_dias}

    # 9. occ_meses payload (chaves normalizadas)
    occ_meses_payload = {}
    for mes_key, mes_data in sorted(occ_data.items()):
        occ_meses_payload[mes_key] = [
            {"nome":emp,"reservas":v["reservas"],"occ_pct":v["occ_pct"],"faturamento":v["faturamento"]}
            for emp, v in mes_data.items()
        ]

    pi = sorted(vm.keys())[0] if vm else ""; pf = sorted(vm.keys())[-1] if vm else ""
    if pi == pf:
        periodo = NM.get(pi[5:7],"") + " " + pi[:4] if pi else ""
    else:
        periodo = NM.get(pi[5:7],"") + " â " + NM.get(pf[5:7],"") + " " + pf[:4] if pi else ""

    return {
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "periodo": periodo,
        "total": gm["total"],
        "tpr_med": gm["tpr_med"],
        "ttf_med": gm["ttf_med"],
        "dias": [{"d":d,"c":c} for d,c in dias_s],
        "semanas": sems,
        "meses": meses_vol,
        "classifs": gm["classifs"],
        "empreendimentos": gm["empreendimentos"],
        "correl": sorted([e for e in gm["empreendimentos"] if e["taxa_contato"] is not None],
                         key=lambda x: -x["taxa_contato"]),
        "clusters": gm["clusters"],
        "prioridades": gm["prioridades"],
        "agentes": gm["agentes"],
        "occ_meses": occ_meses_payload,
        "por_mes": por_mes,
    }

# ââ Uso local ââââââââââââââââââââââââââââââââââââââââââââââââ

if __name__ == "__main__":
    import openpyxl
    print("Carregando planilha...")
    wb = openpyxl.load_workbook(ARQUIVO_ENTRADA)

    sheet_droz = None
    for candidate in ["Atendimentos","Input Droz (Abr-Mai)","Resultado da consulta"]:
        if candidate in wb.sheetnames:
            sheet_droz = wb[candidate]; break
    if not sheet_droz:
        sheet_droz = wb.active
    print("Aba de atendimentos:", sheet_droz.title)

    ws_occ = wb["Reservas e OCC"] if "Reservas e OCC" in wb.sheetnames else None
    print("Aba de OCC:", ws_occ.title if ws_occ else "nÃ£o encontrada")

    rows_droz = list(sheet_droz.iter_rows(min_row=2, values_only=True))
    rows_occ  = list(ws_occ.iter_rows(min_row=2, values_only=True)) if ws_occ else None

    payload = processar(rows_droz, rows_occ)
    print("Tickets Ãºnicos:", payload["total"])
    print("TPR med:", payload["tpr_med"], "| TTF med:", payload["ttf_med"])
    print("OCC meses:", list(payload["occ_meses"].keys()))
    print("Por mÃªs:", {k: payload["por_mes"][k]["total"] for k in payload["por_mes"]})

    with open(ARQUIVO_TEMPLATE,"r",encoding="utf-8") as f:
        html = f.read()
    html = html.replace("__DATA__", json.dumps(payload, ensure_ascii=False))
    with open(ARQUIVO_SAIDA,"w",encoding="utf-8") as f:
        f.write(html)
    print("Dashboard gerado:", ARQUIVO_SAIDA)
