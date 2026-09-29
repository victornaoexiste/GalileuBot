"""
=============================================================
  CADASTRO AUTOMATICO DE EXEMPLARES - SISTEMA GALILEU
=============================================================
COMO USAR:
1. Chrome com depuracao remota aberto e logado no Galileu
2. python cadastrar_exemplares.py
   (para tentar de novo os titulos que ficaram bloqueados como
   "nao encontrados" antes, rode com --retentar:
   python cadastrar_exemplares.py --retentar)
=============================================================
Para cada titulo na biblioteca, cadastra um exemplar com:
  - Disponivel para: TODOS
  - Unidade: CENTRO EDUCATIVO JANUARENSE (ja preenchida)
  - Tombo: automatico
  - Tipo Aquisicao: DOACAO
  - Situacao: DISPONIVEL
=============================================================
"""

import argparse
import time, os, sys, logging, unicodedata, openpyxl
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

BASE = Path(__file__).parent  # Pasta onde o script esta

# ── PASTAS DO PROJETO ──────────────────────────────────────
PASTA_DADOS  = BASE / "dados"
PASTA_LOGS   = BASE / "logs" / "exemplares"

# Cria pastas automaticamente se nao existirem
PASTA_DADOS.mkdir(exist_ok=True)
PASTA_LOGS.mkdir(parents=True, exist_ok=True)

URL_FORMULARIO  = "https://ec2galileu.com.br/admin/titulo-exemplar/formulario"
LOG_OK          = str(PASTA_DADOS / "exemplares_cadastrados.txt")
LOG_NAO_ACHADOS = str(PASTA_DADOS / "exemplares_nao_encontrados.txt")
# IMPORTANTE: o livros_cadastrados.txt de verdade (atualizado pelo
# 1_cadastrar_livros.py) fica na RAIZ do GalileuBot, não dentro de dados/
# — existiam cópias antigas e paradas em dados/ e em "Nova pasta/" que
# causavam essa mesma pasta ser lida por engano, mostrando só uma fração
# dos títulos já cadastrados.
LIVROS_OK       = str(BASE / "livros_cadastrados.txt")
LOG_ERROS       = str(PASTA_LOGS / f"erros_exemplar_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
ESPERA       = 15
PLANILHA        = str(PASTA_DADOS / "planilhaBase.xlsx")

# Mesmo escopo do 1_cadastrar_livros.py: so cria exemplar pros titulos que
# estao na planilha atual (ja filtrada por Eixo Tematico). Isso evita que o
# bot mexa nos ~700 titulos de outras areas que ja foram cadastrados antes.
def titulos_no_escopo():
    if not os.path.exists(PLANILHA):
        log(f"Planilha '{PLANILHA}' nao encontrada — nao foi possivel restringir o escopo!","ERRO")
        sys.exit(1)
    wb = openpyxl.load_workbook(PLANILHA)
    ws = wb.active
    titulos = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            titulos.add(str(row[0]).strip().replace(" PDF","").strip())
    return titulos

logging.basicConfig(
    filename=LOG_ERROS, level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s", encoding="utf-8"
)

def log(msg, tipo="INFO"):
    s = {"INFO":"   ","OK":"OK ","ERRO":"XX ","AVISO":"!! ","INICIO":">> "}
    print(f"{s.get(tipo,'   ')} {msg}")
    getattr(logging, {"OK":"info","ERRO":"error","AVISO":"warning"}.get(tipo,"info"))(msg)

def carregar_nao_achados():
    if not os.path.exists(LOG_NAO_ACHADOS): return set()
    with open(LOG_NAO_ACHADOS, "r", encoding="utf-8") as f:
        return set(l.strip() for l in f if l.strip())

def salvar_nao_achado(nome):
    with open(LOG_NAO_ACHADOS, "a", encoding="utf-8") as f:
        f.write(nome.strip() + "\n")
    log(f"  Registrado em nao_encontrados: {nome[:50]}", "AVISO")

def carregar_ja_feitos():
    """Retorna uma lista (nao um set) — titulos duplicados (mesmo nome,
    exemplares diferentes) precisam ser contados, nao so verificados por
    presenca, senao a 1a ocorrencia feita faz TODAS as outras com o mesmo
    nome serem puladas para sempre apos uma interrupcao/retomada."""
    if not os.path.exists(LOG_OK): return []
    with open(LOG_OK, "r", encoding="utf-8") as f:
        return [l.strip() for l in f if l.strip()]

def salvar_ok(chave):
    with open(LOG_OK, "a", encoding="utf-8") as f:
        f.write(chave.strip() + "\n")

def iniciar_chrome():
    opt = Options()
    opt.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    drv = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opt)
    log("Chrome conectado!", "OK")
    return drv

def clicar(driver, el):
    driver.execute_script("arguments[0].click();", el)

def fechar_modais(driver):
    try:
        driver.execute_script("""
            document.querySelectorAll('.modal-backdrop').forEach(function(b){
                b.parentNode.removeChild(b);
            });
            document.body.classList.remove('modal-open');
            document.body.style.overflow = '';
        """)
        time.sleep(0.4)
    except: pass

# ── PEGA TODOS OS TITULOS DA BIBLIOTECA ─────────────────

# LIVROS_OK ja definido acima

def pegar_titulos(driver):
    """Le os titulos do arquivo livros_cadastrados.txt"""
    caminho_absoluto = Path(LIVROS_OK).resolve()
    log(f"Lendo lista de títulos de: {caminho_absoluto}")

    if not os.path.exists(LIVROS_OK):
        log(f"Arquivo '{LIVROS_OK}' nao encontrado!", "ERRO")
        sys.exit(1)

    with open(LIVROS_OK, "r", encoding="utf-8") as f:
        todos = [l.strip() for l in f if l.strip()]
    log(f"Total de titulos lidos: {len(todos)}", "OK")

    escopo = titulos_no_escopo()
    # Titulos duplicados (mesmo nome, livros diferentes) aparecem 2+ vezes em
    # todos, na ordem em que foram criados no Galileu. O modal de busca do
    # titulo devolve os resultados na mesma ordem (por id crescente), entao
    # guardamos a "ocorrencia" de cada nome pra saber qual resultado clicar
    # — sem isso, todo exemplar duplicado cairia sempre no primeiro titulo.
    contagem = {}
    titulos = []
    for n in todos:
        if n not in escopo:
            continue
        ocorrencia = contagem.get(n, 0)
        contagem[n] = ocorrencia + 1
        # tombo = posicao absoluta nessa lista (estavel entre execucoes,
        # mesmo com --retentar ou apos interrupcao) — evita colidir tombo
        # com o de outro exemplar em runs diferentes
        titulos.append({"nome": n, "id": "", "ocorrencia": ocorrencia, "tombo": len(titulos) + 1})
    log(f"Restrito ao escopo da planilha atual: {len(titulos)} titulos "
        f"({len(todos)-len(titulos)} de outras areas ignorados)", "OK")

    # aviso de sanidade: se existir outra cópia do mesmo nome de arquivo em
    # dados/ com tamanho bem diferente, é sinal de que pode haver confusão
    # de qual arquivo é o "de verdade" (foi exatamente o que aconteceu aqui)
    caminho_alternativo = PASTA_DADOS / "livros_cadastrados.txt"
    if caminho_alternativo.exists() and caminho_alternativo.resolve() != caminho_absoluto:
        tam_usado = caminho_absoluto.stat().st_size
        tam_alt = caminho_alternativo.stat().st_size
        if abs(tam_usado - tam_alt) > 500:
            log(
                f"ATENÇÃO: existe outra cópia de livros_cadastrados.txt em "
                f"'{caminho_alternativo.resolve()}' com tamanho bem diferente "
                f"({tam_alt} bytes vs {tam_usado} bytes em uso). Confira se não são "
                f"duas fontes de dados divergentes.",
                "AVISO",
            )

    return titulos

# ── CADASTRA UM EXEMPLAR ─────────────────────────────────

def cadastrar_exemplar(driver, titulo, n, total):
    nome = titulo["nome"]
    log(f"\n{'='*55}")
    log(f"[{n}/{total}] {nome[:60]}", "INICIO")
    log(f"{'='*55}")

    try:
        driver.get(URL_FORMULARIO)
        time.sleep(3)
        fechar_modais(driver)

        # ── Selecionar Titulo ────────────────────────────
        log(f"  Buscando: {nome[:40]}...")

        # Remove acentos para busca
        nome_ascii_completo = ''.join(
            c for c in unicodedata.normalize('NFD', nome)
            if unicodedata.category(c) != 'Mn'
        )
        if len(nome_ascii_completo) <= 20:
            nome_ascii = nome_ascii_completo.strip(" :-–—,.;")
        else:
            corte = nome_ascii_completo[:20]
            # só recua pra última palavra se realmente cortou uma palavra no
            # meio (ou seja, o próximo caractere ainda faz parte de uma
            # palavra) — evita descartar à toa uma palavra completa que só
            # encostou no limite de 20 caracteres
            if nome_ascii_completo[20:21].isalnum() and " " in corte:
                corte = corte.rsplit(" ", 1)[0]
            nome_ascii = corte.strip(" :-–—,.;")

        # Abre o modal
        driver.execute_script("""
            var btn = document.getElementById('btnLocalizarSelecionarTitulo');
            if(btn){ btn.click(); return; }
            var btns = document.querySelectorAll('label.btn, button');
            for(var i=0;i<btns.length;i++){
                var t = btns[i].getAttribute('title') || btns[i].getAttribute('data-original-title') || '';
                if(t.indexOf('Selecionar T') >= 0){ btns[i].click(); break; }
            }
        """)
        time.sleep(3)

        # Forca modal e campo visiveis
        driver.execute_script("""
            ['dialog-selecionar-titulo'].forEach(function(id){
                var d = document.getElementById(id);
                if(d){ d.style.display='block'; d.style.visibility='visible'; d.classList.add('in'); }
            });
            var c = document.getElementById('tx_titulobiblioteca');
            if(c){
                c.style.display='block'; c.style.visibility='visible';
                c.style.opacity='1'; c.removeAttribute('disabled');
            }
        """)
        time.sleep(0.5)

        def buscar_titulo(termo):
            # Forca campo visivel e interagivel
            driver.execute_script("""
                var c = document.getElementById('tx_titulobiblioteca');
                if(c){
                    c.style.display='block'; c.style.visibility='visible';
                    c.style.opacity='1'; c.removeAttribute('disabled');
                    c.removeAttribute('readonly'); c.value='';
                }
                var d = document.getElementById('dialog-selecionar-titulo');
                if(d){ d.style.display='block'; d.style.visibility='visible'; d.classList.add('in'); }
            """)
            time.sleep(0.5)

            # Usa send_keys real — o sistema precisa de eventos reais de teclado
            try:
                from selenium.webdriver.common.keys import Keys
                campo = driver.find_element(By.ID, "tx_titulobiblioteca")
                driver.execute_script("arguments[0].value='';", campo)
                campo.send_keys(termo)
                time.sleep(0.5)
            except:
                driver.execute_script(f"var c=document.getElementById('tx_titulobiblioteca');if(c){{c.value='{termo}';}}")

            # Clica Localizar
            driver.execute_script("""
                var containers = [
                    document.getElementById('dialog-selecionar-titulo'),
                    document.querySelector('.modal.in'),
                    document.querySelector('.modal[style*="block"]'),
                    document.body
                ];
                for(var k=0;k<containers.length;k++){
                    if(!containers[k]) continue;
                    var btns = containers[k].querySelectorAll('button,label.btn,input[type=button]');
                    for(var i=0;i<btns.length;i++){
                        if((btns[i].textContent||btns[i].value||'').trim().indexOf('Localizar')>=0){
                            btns[i].click(); return;
                        }
                    }
                }
            """)
            time.sleep(3.5)

            elementos = driver.find_elements(By.XPATH,
                "//*[contains(@class,'btnSelecionar') and contains(@class,'alignTitulo')]"
                " | //*[contains(@class,'btnSelecionar')][@id_titulobiblioteca]"
                " | //div[@id='dialog-selecionar-titulo']//*[contains(@class,'btnSelecionar')]"
                " | //div[contains(@class,'modal') and not(contains(@style,'none'))]//*[contains(@class,'btnSelecionar')]"
            )
            log(f"    busca por '{termo}' -> {len(elementos)} resultado(s)")
            return elementos

        resultados = buscar_titulo(nome_ascii)
        if not resultados:
            # Tenta com primeiras 3 palavras
            palavras = nome_ascii.split()
            if len(palavras) >= 3:
                termo3 = ' '.join(palavras[:3])
                log(f"  Tentando 3 palavras '{termo3}'...", "AVISO")
                resultados = buscar_titulo(termo3)
        if not resultados:
            # Tenta com primeiras 2 palavras
            palavras = nome_ascii.split()
            if len(palavras) >= 2:
                termo2 = ' '.join(palavras[:2])
                log(f"  Tentando 2 palavras '{termo2}'...", "AVISO")
                resultados = buscar_titulo(termo2)
        if not resultados:
            # Tenta com primeira palavra
            termo1 = nome_ascii.split()[0] if nome_ascii.split() else nome_ascii[:6]
            log(f"  Tentando 1 palavra '{termo1}'...", "AVISO")
            resultados = buscar_titulo(termo1)
        if not resultados:
            # Tenta sem pontuação (":", "-", etc. podem atrapalhar a busca)
            import re as _re
            sem_pontuacao = _re.sub(r"[^\w\s]", " ", nome_ascii).strip()
            sem_pontuacao = _re.sub(r"\s+", " ", sem_pontuacao)
            if sem_pontuacao and sem_pontuacao != nome_ascii:
                log(f"  Tentando sem pontuação '{sem_pontuacao[:30]}'...", "AVISO")
                resultados = buscar_titulo(sem_pontuacao[:20])

        if not resultados:
            log(f"  Titulo nao encontrado: {nome[:50]}", "ERRO")
            salvar_nao_achado(nome)
            driver.execute_script("""
                var c = document.querySelector('[data-dismiss="modal"],.modal .close,button.close');
                if(c) c.click();
            """)
            return False

        # Titulos duplicados (mesmo nome) retornam varios resultados na
        # ordem de criacao — clica no que corresponde a esta ocorrencia,
        # nao sempre no primeiro (ver comentario em pegar_titulos).
        idx = min(titulo.get("ocorrencia", 0), len(resultados) - 1)
        if idx > 0:
            log(f"  Titulo duplicado: usando resultado {idx+1}/{len(resultados)}", "AVISO")
        driver.execute_script("arguments[0].click();", resultados[idx])
        time.sleep(1.5)
        # Em raras ocasioes o modal de busca de titulo nao fecha sozinho
        # depois do clique em "Selecionar" (bug do site), e fica cobrindo
        # a tela — isso bloqueia o salvamento no final sem nenhuma
        # mensagem de erro visivel. Forca o fechamento por garantia.
        driver.execute_script("""
            var d = document.getElementById('dialog-selecionar-titulo');
            if(d){ d.style.display='none'; d.classList.remove('in'); }
            document.querySelectorAll('.modal-backdrop').forEach(function(b){ b.parentNode.removeChild(b); });
            document.body.classList.remove('modal-open');
        """)
        log("  Titulo selecionado", "OK")

        # ── Disponivel para: TODOS ────────────────────────
        try:
            todos = driver.find_element(By.XPATH,
                "//input[@type='radio'][@value='T' or @value='TODOS']"
                " | //label[normalize-space()='TODOS']/preceding-sibling::input[@type='radio'][1]"
                " | //label[normalize-space()='TODOS']/input"
            )
            clicar(driver, todos)
        except:
            try:
                clicar(driver, driver.find_element(By.XPATH,
                    "//label[normalize-space()='TODOS']"))
            except: pass
        log("  Disponivel: TODOS", "OK")

        # ── Tipo Aquisicao: ACERVO PÚBLICO ONLINE ────────
        try:
            sel = Select(driver.find_element(By.ID,
                "titulo-exemplar-id_tipoaquisicaoexemplarbiblioteca"
            ))
            # Seleciona por indice 1 (sempre ACERVO PÚBLICO ONLINE)
            sel.select_by_index(1)
            log(f"  Tipo Aquisicao: {sel.first_selected_option.text}", "OK")
        except Exception as e:
            log(f"  Tipo Aquisicao erro: {e}", "AVISO")

        # ── Situacao: DISPONÍVEL ──────────────────────────
        try:
            # Usa JS para selecionar DISPONIVEL pelo value ou texto
            driver.execute_script("""
                var sels = document.querySelectorAll('select');
                for(var i=0; i<sels.length; i++){
                    var name = (sels[i].name || sels[i].id || '').toLowerCase();
                    if(name.indexOf('situacao') >= 0 || name.indexOf('situação') >= 0){
                        for(var j=0; j<sels[i].options.length; j++){
                            var txt = sels[i].options[j].text.toUpperCase();
                            if(txt.indexOf('DISPON') >= 0){
                                sels[i].selectedIndex = j;
                                sels[i].dispatchEvent(new Event('change', {bubbles:true}));
                                break;
                            }
                        }
                        break;
                    }
                }
            """)
            time.sleep(0.3)
            # Confirma pelo Select do Selenium
            sel_sit = Select(driver.find_element(By.XPATH,
                "//select[contains(@name,'situacao') or contains(@id,'situacao')]"
            ))
            log(f"  Situacao: {sel_sit.first_selected_option.text}", "OK")
        except Exception as e:
            log(f"  Situacao erro: {e}", "AVISO")

        # ── Tombo (obrigatorio, sem geracao automatica nessa instancia) ──
        try:
            tombo = str(titulo.get("tombo") or n).zfill(6)
            campo_tombo = driver.find_element(By.ID, "titulo-exemplar-tx_numerotombo")
            driver.execute_script("arguments[0].value = arguments[1];", campo_tombo, tombo)
            driver.execute_script("arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                                   "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", campo_tombo)
            log(f"  Tombo: {tombo}", "OK")
        except Exception as e:
            log(f"  Tombo erro: {e}", "AVISO")

        # ── Salvar ────────────────────────────────────────
        fechar_modais(driver)
        driver.execute_script("window.scrollTo(0,0);")
        time.sleep(0.5)

        # Usa o ID exato: btnSalvar
        driver.execute_script("document.getElementById('btnSalvar').click();")
        time.sleep(5)

        # Verifica erro de verdade (campo obrigatorio, mensagem de erro
        # visivel) antes de declarar sucesso — a versao anterior sempre
        # logava "EXEMPLAR CADASTRADO!" mesmo quando nada foi salvo.
        TEXTOS_IGNORAR = ["caso não informado", "configurações", "×", "x"]
        erros = driver.find_elements(By.XPATH,
            "//*[contains(@class,'alert-danger') or contains(@class,'msg-erro')]"
            " | //*[contains(@id,'containerMessage') and (contains(.,'erro') or contains(.,'obrigat') or contains(.,'inválid'))]"
        )
        for e in erros:
            try:
                if not e.is_displayed(): continue
                txt = e.text.strip().lower()
                if not txt or any(ig in txt for ig in TEXTOS_IGNORAR): continue
                log(f"  Erro ao salvar: {e.text.strip()[:120]}", "ERRO")
                return False
            except: continue

        msg_sucesso = None
        try:
            msg = driver.find_element(By.XPATH,
                "//*[contains(@class,'alert-success')]"
                " | //*[contains(@id,'containerMessage') and contains(.,'sucesso')]"
            )
            if msg.is_displayed() and len(msg.text.strip()) > 3:
                msg_sucesso = msg.text.strip()
        except: pass

        if not msg_sucesso:
            log("  Nenhuma confirmacao de sucesso encontrada apos salvar", "ERRO")
            return False

        log(f"  {msg_sucesso[:60]}", "OK")
        log("EXEMPLAR CADASTRADO!", "OK")
        return True

    except Exception as e:
        log(f"Erro: {type(e).__name__}: {str(e)[:150]}", "ERRO")
        logging.exception(f"Exemplar: {nome}")
        return False

# ── MAIN ─────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Cadastro automático de exemplares - Sistema Galileu")
    parser.add_argument(
        "--retentar",
        action="store_true",
        help="tenta de novo os títulos que ficaram marcados como 'não encontrado' em execuções "
        "anteriores, sem precisar editar exemplares_nao_encontrados.txt manualmente",
    )
    args = parser.parse_args()

    print("\n" + "="*55)
    print("  CADASTRO DE EXEMPLARES - SISTEMA GALILEU")
    print("="*55)
    print(f"\n  Log: {LOG_ERROS}\n")
    if args.retentar:
        print("  Modo --retentar ativo: títulos 'não encontrados' antes serão tentados de novo.\n")
    input("Chrome aberto e logado? ENTER para comecar...")

    driver = iniciar_chrome()
    ja_feitos = carregar_ja_feitos()

    titulos = pegar_titulos(driver)
    nao_achados = set() if args.retentar else carregar_nao_achados()

    # Conta quantas vezes cada nome ja foi feito (nao so "esta na lista") —
    # necessario pra titulos duplicados (mesmo nome, exemplares diferentes)
    # nao serem todos pulados so porque um deles ja foi cadastrado.
    from collections import Counter
    feitos_count = Counter(ja_feitos)
    usados = Counter()
    novos = []
    for t in titulos:
        nome = t["nome"]
        if nome in nao_achados:
            continue
        if usados[nome] < feitos_count.get(nome, 0):
            usados[nome] += 1
            continue
        novos.append(t)

    if nao_achados:
        log(f"{len(nao_achados)} titulos ignorados (nao encontrados anteriormente):")
        for n in sorted(nao_achados):
            log(f"  - {n[:60]}", "AVISO")
    elif args.retentar:
        log("Ignorando a lista de 'não encontrados' anteriores por causa do --retentar.")

    log(f"{len(novos)} para cadastrar | {len(ja_feitos)} ja feitos")

    if not novos:
        log("Todos os exemplares ja foram cadastrados!", "OK"); return

    ok = falha = 0
    try:
        for i, titulo in enumerate(novos, 1):
            if cadastrar_exemplar(driver, titulo, i, len(novos)):
                ok += 1
                salvar_ok(titulo["nome"])
            else:
                falha += 1
            time.sleep(2)
    except KeyboardInterrupt:
        log("Interrompido.", "AVISO")
    finally:
        print("\n" + "="*55)
        log(f"Sucesso : {ok}", "OK")
        if falha: log(f"Com erro: {falha}", "AVISO")
        log(f"Log: {LOG_ERROS}")
        input("\nENTER para encerrar...")

if __name__ == "__main__":
    main()
