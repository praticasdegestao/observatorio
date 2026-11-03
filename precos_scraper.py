#!/usr/bin/env python3
"""
Atualiza o painel de preços do Radar de Mercado (precos.json).

Fontes
  - HF Brasil (Hortifruti/Cepea): banana, citros e mamão  -> bloco "hf"
  - Agrolink (cotações das Ceasas): mandioca, banana, abacaxi, mamão, laranja,
    limão, tangerina e maracujá                               -> bloco "ceasa"
  - Notícias Agrícolas: preços semanais da fécula de mandioca (dados Cepea/Esalq) -> bloco "hf"
  - Cepea (indicador da mandioca): o site proíbe robôs (robots.txt), por isso
    NÃO é raspado. Baixe a planilha do indicador no site do Cepea e salve como
    cepea_mandioca.xls (ou .xlsx) nesta pasta: o script a lê como arquivo local.

Uso:  pip install requests beautifulsoup4 lxml pandas openpyxl xlrd
      python precos_scraper.py            # grava precos.json ao lado
Publique precos.json na mesma pasta do index.html (ex.: GitHub Pages).

Boas práticas: o script consulta o robots.txt de cada site, identifica-se no
User-Agent, espera alguns segundos entre requisições e, se uma fonte falhar,
mantém os dados anteriores dessa fonte (não apaga o que já existia).
Os layouts dos sites mudam; se um bloco vier vazio, ajuste o parser indicado.
"""
import datetime as dt, io, json, os, re, sys, time, urllib.robotparser
from urllib.parse import urlparse
import requests
import pandas as pd
from bs4 import BeautifulSoup

UA = "PainelProducaoAgricola/1.0 (+uso institucional; contato no README)"
PAUSA = 4
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "precos.json")
HOJE = dt.date.today()

HF = [("Banana", "https://www.hfbrasil.org.br/br/estatistica/banana.aspx"),
      ("Citros", "https://www.hfbrasil.org.br/br/estatistica/citros.aspx"),
      ("Mamão", "https://www.hfbrasil.org.br/br/estatistica/mamao.aspx")]
CEASA = [("Mandioca", "Mandioca / aipim", "kg", "https://www.agrolink.com.br/cotacoes/ceasa/diversos/mandioca/mandioca-aipim-1kg"),
         ("Banana", "Banana nanica", "kg", "https://www.agrolink.com.br/cotacoes/ceasa/frutas/banana"),
         ("Abacaxi", "Abacaxi", "unidade", "https://www.agrolink.com.br/cotacoes/ceasa/frutas/abacaxi/abacaxi-1un"),
         ("Mamão", None, "kg", "https://www.agrolink.com.br/cotacoes/ceasa/frutas/mamao"),
         ("Citros", "Laranja Pêra", "kg", "https://www.agrolink.com.br/cotacoes/ceasa/frutas/laranja/laranja-pera-1kg"),
         ("Citros", "Limão Taiti", "kg", "https://www.agrolink.com.br/cotacoes/ceasa/frutas/limao/limao-taiti-1kg"),
         ("Citros", "Tangerina", "kg", "https://www.agrolink.com.br/cotacoes/ceasa/frutas/tangerina/tangerina-1kg"),
         ("Maracujá", "Maracujá azedo", "kg", "https://www.agrolink.com.br/cotacoes/ceasa/frutas/maracuja/maracuja-azedo-1kg")]
CEPEA_URL = "https://www.cepea.org.br/br/indicador/mandioca.aspx"
NA_FECULA = "https://www.noticiasagricolas.com.br/cotacoes/mandioca/precos-da-fecula-de-mandioca"

_robots = {}
def permitido(url):
    base = "{0.scheme}://{0.netloc}".format(urlparse(url))
    if base not in _robots:
        rp = urllib.robotparser.RobotFileParser(base + "/robots.txt")
        try: rp.read()
        except Exception: rp = None
        _robots[base] = rp
    rp = _robots[base]
    return True if rp is None else rp.can_fetch(UA, url)

def baixar(url):
    if not permitido(url):
        raise PermissionError(f"robots.txt não permite {url}")
    r = requests.get(url, headers={"User-Agent": UA, "Accept-Language": "pt-BR"}, timeout=40)
    r.raise_for_status(); time.sleep(PAUSA)
    return r.text

def num(s):
    if isinstance(s, (int, float)): return None if pd.isna(s) else float(s)
    if s is None or str(s).strip() in ("", "-", "- - -", "—", "nan"): return None
    t = str(s).replace("R$", "").strip()
    if "," in t: t = t.replace(".", "").replace(",", ".")
    try: return float(t)
    except ValueError: return None

def data_iso(s, ano=HOJE.year):
    m = re.search(r"(\d{1,2})/(\d{1,2}|[a-zç]{3})(?:/(\d{4}))?", str(s), re.I)
    if not m: return None
    meses = dict(jan=1, fev=2, mar=3, abr=4, mai=5, jun=6, jul=7, ago=8, set=9, out=10, nov=11, dez=12)
    d, mm, y = int(m.group(1)), m.group(2), m.group(3)
    mm = int(mm) if mm.isdigit() else meses.get(mm.lower()[:3])
    y = int(y) if y else (ano - 1 if mm > HOJE.month + 1 else ano)   # tabelas sem ano: virada de ano
    return f"{y:04d}-{mm:02d}-{d:02d}"

# ---------------- HF Brasil: tabela "Produto | Região | Unidade | datas..." ----------------
def parse_hf(cult, url):
    html = baixar(url)
    for t in pd.read_html(io.StringIO(html), decimal=",", thousands=".", flavor="lxml"):
        cols = [str(c).strip() for c in t.columns]
        if len(cols) >= 4 and cols[0].lower().startswith("produto") and "regi" in cols[1].lower():
            dcols = [c for c in cols[3:] if data_iso(c)]
            rows = []
            for _, r in t.iterrows():
                v = [num(r[c]) for c in dcols]
                if all(x is None for x in v): continue
                rows.append({"produto": str(r[cols[0]]).strip(), "regiao": str(r[cols[1]]).strip(), "unidade": str(r[cols[2]]).strip().replace("Kg", "kg").replace("quilograma", "kg"), "v": v})
            if rows:
                return {"cult": cult, "fonte": "HF Brasil (Hortifruti/Cepea)", "url": url, "cols": [data_iso(c) for c in dcols], "rows": rows}
    raise ValueError("tabela de preços não encontrada (layout mudou?)")

# ---------------- Agrolink: lista por praça ou tabela com histórico ----------------
UF_RE = re.compile(r"\b(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b")
def limpa_praca(s):
    s = re.sub(r"\s*-\s*[A-Z]{2}\s*·\s*", " ", s); s = re.sub(r"CEASA-[A-Z]{2}", "CEASA", s); s = re.sub(r"AMA-[A-Z]{2}", "AMA", s)
    return re.sub(r"\s+", " ", s).strip()

def parse_agrolink(cult, produto, unidade, url):
    html = baixar(url)
    rows = []
    # 1) páginas com tabela HTML (Produto | Ceasa | Preço | Data)
    try:
        for t in pd.read_html(io.StringIO(html), decimal=",", thousands=".", flavor="lxml"):
            cols = [str(c).lower() for c in t.columns]
            ip = next((i for i, c in enumerate(cols) if "preço" in c or "preco" in c), None)
            idt = next((i for i, c in enumerate(cols) if "data" in c), None)
            ic = next((i for i, c in enumerate(cols) if "ceasa" in c or "local" in c or "praça" in c), None)
            ipr = next((i for i, c in enumerate(cols) if "produto" in c), None)
            if ip is None or idt is None or ic is None: continue
            for _, r in t.iterrows():
                p, d = num(r.iloc[ip]), data_iso(r.iloc[idt])
                local = str(r.iloc[ic]); uf = UF_RE.search(local)
                if p is None or not d or not uf: continue
                prod = produto or re.sub(r"\s*1\s*kg.*$", "", str(r.iloc[ipr]), flags=re.I).strip()
                rows.append({"produto": prod, "praca": re.sub(r"\s*\([A-Z]{2}\)\s*", " ", limpa_praca(local)).strip(), "uf": uf.group(1), "data": d, "freq": None, "preco": p, "var": None})
    except (ValueError, ImportError):
        pass
    hist = bool(rows)
    # 2) páginas em lista (praça · data · frequência · preço · variação)
    if not rows:
        soup = BeautifulSoup(html, "lxml")
        for el in soup.find_all(string=re.compile(r"R\$\s*\d")):
            bloco = el.find_parent(["li", "tr", "div"])
            if not bloco: continue
            txt = " ".join(bloco.get_text(" ", strip=True).split())
            m = re.search(r"((?:CEASA|CEAGESP|AMA|CEASAMINAS)[^·|]*?)\s*[·|-]?\s*([A-ZÀ-Úa-zà-ú .']+?)\s+(\d{2}/\d{2})\s+(Diário|Semanal|Mensal)\s+R\$\s*([\d.,]+)\s*([▲▼][\s\d,]+%|estável)?", txt)
            if not m: continue
            uf = UF_RE.search(m.group(1)) or UF_RE.search(txt)
            v = m.group(6)
            var = 0.0 if v == "estável" else (None if not v else (-1 if "▼" in v else 1) * num(v.replace("▲", "").replace("▼", "").replace("%", "")))
            rows.append({"produto": produto, "praca": limpa_praca(f"{m.group(1).split('-')[0].strip()} {m.group(2).strip()}"), "uf": uf.group(1) if uf else "",
                         "data": data_iso(m.group(3)), "freq": m.group(4), "preco": num(m.group(5)), "var": var})
        # remove duplicatas (o mesmo bloco aparece em divs aninhadas)
        rows = list({(r["praca"], r["data"], r["preco"]): r for r in rows}.values())
    if not rows: raise ValueError("nenhuma cotação reconhecida (layout mudou?)")
    blocos = {}
    for r in rows: blocos.setdefault(r["produto"], []).append(r)
    return [{"cult": cult, "produto": p, "unidade": unidade, "url": url, "fonte": "Agrolink (cotações das Ceasas)", "total": None, "hist": hist, "rows": rr} for p, rr in blocos.items()]

# ---------------- Notícias Agrícolas: fécula de mandioca (dados Cepea/Esalq), semanal ----------------
def parse_na_fecula():
    html = baixar(NA_FECULA)
    sem = {}   # semana (data final ISO) -> {"periodo": texto, regiao: valor}
    for t in pd.read_html(io.StringIO(html), decimal=",", thousands=".", flavor="lxml"):
        cols = [str(c).strip().lower() for c in t.columns]
        if not any("regi" in c for c in cols) or not any("valor" in c for c in cols): continue
        idt = next(i for i, c in enumerate(cols) if "data" in c) if any("data" in c for c in cols) else None
        ir = next(i for i, c in enumerate(cols) if "regi" in c); iv = next(i for i, c in enumerate(cols) if "valor" in c)
        for _, r in t.iterrows():
            per = str(r.iloc[idt]) if idt is not None else ""
            m = re.findall(r"(\d{1,2}/\d{1,2}/\d{4})", per)
            if not m: continue
            fim = data_iso(m[-1]); v = num(r.iloc[iv])
            if v is None: continue
            s = sem.setdefault(fim, {"periodo": per.strip()}); s[str(r.iloc[ir]).strip()] = v
    if not sem: raise ValueError("tabela da fécula não encontrada (layout mudou?)")
    cols = sorted(sem); regs = sorted({k for s in sem.values() for k in s if k != "periodo"})
    return {"cult": "Mandioca", "fonte": "Notícias Agrícolas (dados Cepea/Esalq)", "titulo": "Preços da fécula de mandioca", "defMode": "raw", "url": NA_FECULA,
            "cols": cols, "periodo": [sem[c]["periodo"] for c in cols],
            "rows": [{"produto": "Fécula de mandioca", "regiao": g, "unidade": "t", "v": [sem[c].get(g) for c in cols]} for g in regs]}

# ---------------- Cepea: só arquivo local baixado pelo usuário ----------------
def cepea_local():
    for nome in ("cepea_mandioca.xlsx", "cepea_mandioca.xls"):
        f = os.path.join(HERE, nome)
        if not os.path.exists(f): continue
        bruto = pd.read_excel(f, header=None)
        h = next(i for i in range(min(15, len(bruto))) if str(bruto.iloc[i, 0]).strip().lower().startswith("data"))
        df = pd.read_excel(f, header=h).dropna(how="all")
        df = df[df.iloc[:, 0].notna()].tail(8)
        datas = [data_iso(pd.to_datetime(x, dayfirst=True).strftime("%d/%m/%Y")) for x in df.iloc[:, 0]]
        rows = [{"produto": str(c).strip(), "regiao": "Indicador Cepea", "unidade": "ver planilha", "v": [num(x) if isinstance(x, str) else (None if pd.isna(x) else float(x)) for x in df[c]]} for c in df.columns[1:]]
        return {"cult": "Mandioca", "fonte": "Cepea/Esalq-USP", "titulo": "Indicador da mandioca", "url": CEPEA_URL, "status": "ok", "cols": datas, "rows": rows}
    return {"cult": "Mandioca", "fonte": "Cepea/Esalq-USP", "titulo": "Indicador da mandioca (raiz e fécula)", "url": CEPEA_URL, "status": "bloqueado",
            "msg": "O painel não coleta dados do Cepea (o site proíbe robôs). Abaixo, a própria página do indicador é exibida diretamente do site do Cepea pelo seu navegador, como uma visita normal: nada é copiado nem armazenado pelo painel."}

def main():
    antigo = json.load(open(OUT, encoding="utf-8")) if os.path.exists(OUT) else {}
    out = {"coletado": HOJE.isoformat(), "src": "precos.json", "cepea": [cepea_local()], "hf": [], "ceasa": [], "obs": [], "erros": []}
    try: out["hf"].append(parse_na_fecula()); print("ok  Notícias Agrícolas · fécula")
    except Exception as e:
        out["erros"].append(f"Notícias Agrícolas fécula: {e}"); print("ERRO NA fécula", e, file=sys.stderr)
        out["hf"] += [b for b in antigo.get("hf", []) if b["url"] == NA_FECULA]
    for cult, url in HF:
        try: out["hf"].append(parse_hf(cult, url)); print("ok  HF", cult)
        except Exception as e:
            out["erros"].append(f"HF {cult}: {e}"); print("ERRO HF", cult, e, file=sys.stderr)
            out["hf"] += [b for b in antigo.get("hf", []) if b["cult"] == cult]
    for cult, prod, un, url in CEASA:
        try: out["ceasa"] += parse_agrolink(cult, prod, un, url); print("ok  Agrolink", prod or cult)
        except Exception as e:
            out["erros"].append(f"Agrolink {prod or cult}: {e}"); print("ERRO Agrolink", prod or cult, e, file=sys.stderr)
            out["ceasa"] += [b for b in antigo.get("ceasa", []) if b["url"] == url]
    # histórico: junta as coletas anteriores (últimos 180 dias) para os gráficos de linha
    corte = (HOJE - dt.timedelta(days=180)).isoformat()
    for b in out["ceasa"]:
        velho = next((o for o in antigo.get("ceasa", []) if o["url"] == b["url"] and o["produto"] == b["produto"]), None)
        if velho and velho is not b:
            vistos = {(r["praca"], r["uf"], r["data"]) for r in b["rows"]}
            b["rows"] += [r for r in velho["rows"] if (r["praca"], r["uf"], r["data"]) not in vistos and r["data"] >= corte]
            b["hist"] = True
    for b in out["hf"]:
        velho = next((o for o in antigo.get("hf", []) if o["url"] == b["url"]), None)
        if velho and velho is not b:
            cols = sorted(set(b["cols"]) | {c for c in velho["cols"] if c >= corte})
            def serie(bl, r):
                m = dict(zip(bl["cols"], r["v"]))
                return m
            idx = {}
            for bl in (velho, b):
                for r in bl["rows"]:
                    k = (r["produto"], r["regiao"], r["unidade"])
                    idx.setdefault(k, {}).update({c: v for c, v in serie(bl, r).items() if v is not None})
            if "periodo" in b or "periodo" in velho:
                per = dict(zip(velho.get("cols", []), velho.get("periodo", []))); per.update(zip(b["cols"], b.get("periodo", [])))
                b["periodo"] = [per.get(c, c) for c in cols]
            b["cols"] = cols
            b["rows"] = [{"produto": k[0], "regiao": k[1], "unidade": k[2], "v": [m.get(c) for c in cols]} for k, m in idx.items()]
    out["obs"] = ["Agrolink informa que os preços das Ceasas são referenciais e não auditados."] + ([f"Fontes com falha nesta coleta (mantidos os dados anteriores): {'; '.join(out['erros'])}."] if out["erros"] else [])
    json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("gravado", OUT)

if __name__ == "__main__":
    main()
