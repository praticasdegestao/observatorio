#!/usr/bin/env python3
"""
Radar de informações - mandioca, banana, abacaxi, mamão e citros.

Coleta notícias recentes da mídia especializada e de instituições, classifica por
cultura e tema com palavras-chave e grava radar.json no formato lido pela aba
"📡 Radar" do dashboard (o HTML lê radar.json se estiver na mesma pasta, servido por http/https).

Uso:
    python radar_scraper.py                 # últimos 60 dias, grava radar.json
    python radar_scraper.py --dias 30 --saida radar.json

Somente biblioteca padrão do Python (3.9+). A coleta usa o RSS de busca do Google
Notícias restrito a cada site (operador site:), o que evita raspar o HTML dos
portais (vários proíbem robôs) e respeita os termos de uso de cada um. Itens já
existentes no radar.json marcados como curados (sem "auto": true) são preservados.
"""
import argparse, datetime as dt, email.utils, html, json, os, re, time, unicodedata
import urllib.parse, urllib.request
import xml.etree.ElementTree as ET

FONTES = {  # veículo -> domínio
    "Globo Rural": "globorural.globo.com",
    "Band Agro": "band.com.br/agro",
    "Notícias Agrícolas": "noticiasagricolas.com.br",
    "Canal Rural": "canalrural.com.br",
    "Revista Cultivar": "revistacultivar.com.br",
    "Agrolink": "agrolink.com.br",
    "Embrapa": "embrapa.br",
    "CNA": "cnabrasil.org.br",
    "Cepea": "cepea.org.br",
    "HF Brasil (Cepea)": "hfbrasil.org.br",
    "Conab": "gov.br/conab",
    "MAPA": "gov.br/agricultura",
    "Fundecitrus": "fundecitrus.com.br",
}
CULTURAS = {  # cultura -> (termos de busca, regex de classificação)
    "Mandioca": ("mandioca OR fécula OR macaxeira OR aipim", r"mandioc|fecula|macaxeira|aipim|farinha de mandioca"),
    "Banana": ("banana OR bananicultura", r"banan"),
    "Abacaxi": ("abacaxi", r"abacaxi|ananas"),
    "Mamão": ("mamão OR papaya", r"mamao|papaya"),
    "Citros": ("laranja OR citros OR citricultura OR greening OR limão OR tangerina", r"laranj|citro|citric|greening|\bhlb\b|limao|tangerin|suco de laranja"),
}
TEMAS = {
    "Preços": r"preco|cotac|indicador|r\$ ?\d|valoriz|desvaloriz|recua|sobe|cai\b|alta de|queda de",
    "Mercado": r"export|import|demanda|consumo|tarifa|mercado|comercializ|embarque|varejo|atacado|ceasa",
    "Produção": r"safra|produc|colheita|area plantada|area colhida|produtividade|estimativa|plantio",
    "P&D": r"pesquis|cultivar|variedade|embrapa|estudo|melhoramento|genetic|inovac|lancament|tecnologia",
    "Agricultura digital": r"digital|drone|sensor|inteligencia artificial|\bia\b|aplicativo|satelit|sensoriamento|software|agtech|startup|robo|machine learning",
    "Clima": r"chuva|seca|clima|el nino|la nina|geada|calor|estiagem|temperatura|hidric|inmet",
    "Sanidade": r"praga|doenc|greening|fusari|antracnose|virus|fungo|psilideo|cochonilha|mosca|sanitari|fitossanit",
    "Políticas": r"mapa\b|ministerio|governo|\blei\b|programa|credito|plano safra|zoneamento|zarc|portaria|politica|pronaf",
}
UA = "Mozilla/5.0 (radar-agro; uso de pesquisa)"


def norm(s):
    return unicodedata.normalize("NFD", s or "").encode("ascii", "ignore").decode().lower()


def busca_rss(consulta):
    url = "https://news.google.com/rss/search?" + urllib.parse.urlencode({"q": consulta, "hl": "pt-BR", "gl": "BR", "ceid": "BR:pt-419"})
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=30) as r:
        return ET.fromstring(r.read())


def limpa(txt):
    txt = re.sub(r"<[^>]+>", " ", html.unescape(txt or ""))
    return re.sub(r"\s+", " ", txt).strip()


def coleta(dias):
    itens, vistos = [], set()
    for veiculo, dom in FONTES.items():
        for cult, (termos, _) in CULTURAS.items():
            q = f"({termos}) site:{dom} when:{dias}d"
            try:
                root = busca_rss(q)
            except Exception as e:  # rede, limite de requisições etc.
                print(f"[aviso] {veiculo}/{cult}: {e}")
                continue
            for it in root.iter("item"):
                titulo = limpa(it.findtext("title"))
                src = it.find("source")
                if src is not None and titulo.endswith(" - " + (src.text or "")):
                    titulo = titulo[: -len(" - " + src.text)]
                chave = norm(re.sub(r"\W+", " ", titulo))[:120]
                if not titulo or chave in vistos:
                    continue
                try:
                    data = email.utils.parsedate_to_datetime(it.findtext("pubDate")).date().isoformat()
                except Exception:
                    data = dt.date.today().isoformat()
                resumo = limpa(it.findtext("description"))
                if resumo.startswith(titulo):
                    resumo = ""
                texto = norm(titulo + " " + resumo)
                culturas = [c for c, (_, rx) in CULTURAS.items() if re.search(rx, texto)] or [cult]
                if not any(re.search(CULTURAS[c][1], texto) for c in culturas):
                    continue  # o título não fala da cultura: descarta falso positivo
                temas = [t for t, rx in TEMAS.items() if re.search(rx, texto)] or ["Mercado"]
                vistos.add(chave)
                itens.append({"data": data, "titulo": titulo, "veiculo": veiculo, "orgao": "", "culturas": culturas,
                              "temas": temas, "resumo": resumo[:400], "url": it.findtext("link"), "auto": True})
            time.sleep(1.0)  # gentileza com o servidor
    return itens


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dias", type=int, default=60)
    ap.add_argument("--saida", default="radar.json")
    ap.add_argument("--manter-dias", type=int, default=365, help="descarta itens mais antigos que isso")
    a = ap.parse_args()
    antigo = {"itens": []}
    if os.path.exists(a.saida):
        with open(a.saida, encoding="utf-8") as f:
            antigo = json.load(f)
    curados = [x for x in antigo.get("itens", []) if not x.get("auto")]
    novos = coleta(a.dias)
    chaves = {norm(re.sub(r"\W+", " ", x["titulo"]))[:120] for x in curados}
    auto_antigos = [x for x in antigo.get("itens", []) if x.get("auto")]
    todos, vistos = list(curados), set(chaves)
    for x in novos + auto_antigos:
        k = norm(re.sub(r"\W+", " ", x["titulo"]))[:120]
        if k not in vistos:
            vistos.add(k); todos.append(x)
    corte = (dt.date.today() - dt.timedelta(days=a.manter_dias)).isoformat()
    todos = sorted([x for x in todos if x["data"] >= corte], key=lambda x: x["data"], reverse=True)
    saida = {"atualizado": dt.date.today().isoformat(),
             "metodo": "Curadoria manual + coleta automática (RSS de busca do Google Notícias por site e cultura; tema por palavras-chave).",
             "itens": todos}
    if antigo.get("sinteses"):
        saida["sinteses"] = antigo["sinteses"]
    with open(a.saida, "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=1)
    print(f"{len(novos)} coletados agora · {len(todos)} no radar · gravado em {a.saida}")


if __name__ == "__main__":
    main()
