"""
Gera uma página HTML com um botão de WhatsApp para cada cliente da planilha,
já com uma mensagem personalizada convidando a pessoa a votar.

Quem vota é o próprio cliente: a mensagem só leva o link da votação.

Uso:
    pip install openpyxl
    python gerar_links_whatsapp.py Clientes_ADVBOX.xlsx [saida.html]

Depois abra o HTML no navegador (com o WhatsApp Web logado), clique em
"Enviar" em cada cliente, confira a mensagem e envie. A página marca
quem já recebeu (fica salvo no navegador).
"""
import html
import json
import re
import sys
import urllib.parse

import openpyxl

LINK_VOTACAO = "https://votacao.escritoriosdigitais.com.br/votar/970b2/laize-da-silva-praxedes"

MENSAGEM = (
    "Olá, {nome}! Tudo bem? Aqui é a Dra. Laize Praxedes.\n\n"
    "Estou concorrendo a um prêmio do Escritórios Digitais e ficaria muito feliz "
    "com o seu voto. É rapidinho, basta acessar o link abaixo:\n\n"
    f"{LINK_VOTACAO}\n\n"
    "Muito obrigada pelo carinho e pela confiança! 💛"
)


def normalizar_telefone(texto):
    """Retorna o primeiro celular válido no formato 55DDXXXXXXXXX, ou None."""
    for parte in re.split(r"[/;,]", str(texto or "")):
        d = re.sub(r"\D", "", parte)
        if d.startswith("55") and len(d) in (12, 13):
            d = d[2:]
        if len(d) == 10 and d[2] in "6789":  # celular antigo sem o 9
            d = d[:2] + "9" + d[2:]
        if len(d) == 11 and d[2] == "9":
            return "55" + d
    return None


def primeiro_nome(nome):
    p = str(nome or "").strip().split()
    return p[0].capitalize() if p else ""


def main(arquivo, saida):
    ws = openpyxl.load_workbook(arquivo, read_only=True).active
    linhas = list(ws.iter_rows(min_row=2, values_only=True))

    contatos, vistos = [], set()
    pulados = {"menor": 0, "sem_celular": 0, "duplicado": 0}
    for nome, _cpf, tel, outros, origem, *_ in linhas:
        if "menor" in str(origem or "").lower():
            pulados["menor"] += 1
            continue
        fone = normalizar_telefone(tel) or normalizar_telefone(outros)
        if not fone:
            pulados["sem_celular"] += 1
            continue
        if fone in vistos:
            pulados["duplicado"] += 1
            continue
        vistos.add(fone)
        msg = MENSAGEM.format(nome=primeiro_nome(nome))
        url = f"https://wa.me/{fone}?text={urllib.parse.quote(msg)}"
        contatos.append({"nome": str(nome).strip(), "fone": fone, "url": url})

    contatos.sort(key=lambda c: c["nome"].lower())
    linhas_html = "\n".join(
        f'<tr data-id="{c["fone"]}"><td><input type="checkbox"></td>'
        f'<td>{html.escape(c["nome"])}</td>'
        f'<td>+{c["fone"][:2]} ({c["fone"][2:4]}) {c["fone"][4:9]}-{c["fone"][9:]}</td>'
        f'<td><a href="{html.escape(c["url"])}" target="_blank">Enviar</a></td></tr>'
        for c in contatos
    )

    pagina = f"""<!doctype html><html lang="pt-BR"><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Convites de votação</title>
<style>
body{{font-family:system-ui,sans-serif;margin:16px;background:#fff;color:#222}}
table{{border-collapse:collapse;width:100%}}td,th{{padding:6px 8px;border-bottom:1px solid #ddd;text-align:left}}
a{{background:#25d366;color:#fff;padding:4px 12px;border-radius:6px;text-decoration:none}}
tr.ok{{opacity:.45}} #busca{{padding:6px;width:100%;max-width:320px;margin:8px 0}}
</style>
<h1>Convites de votação</h1>
<p>{len(contatos)} contatos · <span id="cont">0</span> enviados.
Pulados: {pulados["menor"]} menores, {pulados["sem_celular"]} sem celular, {pulados["duplicado"]} telefones repetidos.</p>
<input id="busca" placeholder="Buscar nome...">
<table><tr><th>✓</th><th>Nome</th><th>WhatsApp</th><th></th></tr>
{linhas_html}
</table>
<script>
const K='convites_enviados';let s={{}};try{{s=JSON.parse(localStorage.getItem(K)||'{{}}')}}catch(e){{}}
const salvar=()=>{{try{{localStorage.setItem(K,JSON.stringify(s))}}catch(e){{}};
  document.getElementById('cont').textContent=Object.values(s).filter(Boolean).length}};
document.querySelectorAll('tr[data-id]').forEach(tr=>{{
  const id=tr.dataset.id,cb=tr.querySelector('input');
  const marca=v=>{{cb.checked=v;tr.classList.toggle('ok',v);s[id]=v;salvar()}};
  if(s[id])marca(true);
  cb.onchange=()=>marca(cb.checked);
  tr.querySelector('a').addEventListener('click',()=>marca(true));
}});
document.getElementById('busca').oninput=e=>{{const q=e.target.value.toLowerCase();
  document.querySelectorAll('tr[data-id]').forEach(tr=>tr.style.display=
    tr.children[1].textContent.toLowerCase().includes(q)?'':'none')}};
salvar();
</script></html>"""

    with open(saida, "w", encoding="utf-8") as f:
        f.write(pagina)
    print(f"{len(contatos)} contatos gerados em {saida}")
    print("Pulados:", json.dumps(pulados, ensure_ascii=False))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "convites_whatsapp.html")
