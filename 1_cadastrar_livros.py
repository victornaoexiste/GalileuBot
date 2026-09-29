"""
=============================================================
  CADASTRO AUTOMATICO DE LIVROS - SISTEMA GALILEU
  Colegio Politecnico Dom Luciano
=============================================================
COMO USAR:
1. Abra o Chrome com depuracao remota:
   google-chrome --remote-debugging-port=9222 --user-data-dir="$(pwd)/ChromeDebug"
2. Faca login no Galileu
3. Confira dados/planilhaBase.xlsx
4. Rode: venv/bin/python 1_cadastrar_livros.py
=============================================================
"""

import time, os, sys, openpyxl, logging
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

# ── CONFIGURACOES ─────────────────────────────────────────
# Pasta raiz = onde esta o script
BASE = Path(__file__).parent

PASTA_PDFS  = BASE / "pdfs"
PASTA_DADOS = BASE / "dados"
PASTA_LOGS  = BASE / "logs" / "livros"
PASTA_PDFS.mkdir(exist_ok=True)
PASTA_DADOS.mkdir(exist_ok=True)
PASTA_LOGS.mkdir(parents=True, exist_ok=True)

PASTA_LIVROS = str(PASTA_PDFS)
PLANILHA     = str(PASTA_DADOS / "planilhaBase.xlsx")
URL_CADASTRO = "https://ec2galileu.com.br/admin/titulo-biblioteca/formulario"
LOG_OK       = str(BASE / "livros_cadastrados.txt")
LOG_ERROS    = str(PASTA_LOGS / f"erros_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
ESPERA       = 15

# Só cadastra livros cujo "Eixo Temático" (coluna I da planilha) esteja aqui.
# Para incluir outro eixo, basta adicionar o texto exato da coluna I.
EIXOS_DESEJADOS = {
    "Enfermagem / Saúde",
    "Administração / Gestão",
    "Variados",
}
# ──────────────────────────────────────────────────────────

logging.basicConfig(
    filename=LOG_ERROS, level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s", encoding="utf-8"
)

def log(msg, tipo="INFO"):
    s = {"INFO":"   ","OK":"OK ","ERRO":"XX ","AVISO":"!! ","INICIO":">> "}
    print(f"{s.get(tipo,'   ')} {msg}")
    getattr(logging,{"OK":"info","ERRO":"error","AVISO":"warning"}.get(tipo,"info"))(msg)

def carregar_cadastrados():
    if not os.path.exists(LOG_OK): return set()
    with open(LOG_OK,"r",encoding="utf-8") as f:
        return set(l.strip() for l in f if l.strip())

def salvar_ok(titulo):
    with open(LOG_OK,"a",encoding="utf-8") as f:
        f.write(" ".join(titulo.split())+"\n")

def normalizar_titulo(s):
    """Colapsa quebras de linha/espacos internos em um so espaco.
    Sem isso, um titulo com Alt+Enter na planilha vira varias 'linhas'
    diferentes no log de cadastrados e o bot recadastra o mesmo livro
    a cada execucao (ja aconteceu: um titulo foi parar 7x no Galileu)."""
    return " ".join(str(s).split())

def carregar_planilha():
    if not os.path.exists(PLANILHA):
        log(f"Planilha '{PLANILHA}' nao encontrada!","ERRO"); sys.exit(1)
    wb = openpyxl.load_workbook(PLANILHA)
    ws = wb.active
    livros = []
    ignorados = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        titulo = normalizar_titulo(row[0]) if row[0] else ""
        if not titulo or titulo=="None": continue
        def v(i): return str(row[i]).strip().replace(" PDF","").strip() if row[i] and len(row) > i else ""
        eixo = v(8)
        if EIXOS_DESEJADOS and eixo not in EIXOS_DESEJADOS:
            ignorados += 1
            continue
        livros.append({
            "titulo":  titulo,
            "tipo":    v(1), "classif": v(2), "editora": v(3),
            "idioma":  v(4), "autores": v(5), "genero":  v(6), "arquivo": v(7),
            "eixo":    eixo,
        })
    log(f"{len(livros)} livros na planilha (eixos: {', '.join(sorted(EIXOS_DESEJADOS))}); {ignorados} ignorados por eixo.")
    return livros

def iniciar_chrome():
    log("Conectando ao Chrome...","INICIO")
    opt = Options()
    opt.add_experimental_option("debuggerAddress","127.0.0.1:9222")
    try:
        drv = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opt)
        log("Chrome conectado!","OK"); return drv
    except Exception as e:
        log(f"Erro: {e}","ERRO"); sys.exit(1)

def W(driver, by, val, t=None):
    return WebDriverWait(driver, t or ESPERA).until(
        EC.presence_of_element_located((by, val)))

def WV(driver, by, val, t=None):
    """Espera elemento ficar visivel"""
    return WebDriverWait(driver, t or ESPERA).until(
        EC.visibility_of_element_located((by, val)))

def WC(driver, by, val, t=None):
    return WebDriverWait(driver, t or ESPERA).until(
        EC.element_to_be_clickable((by, val)))

def clicar(driver, el):
    driver.execute_script("arguments[0].click();", el)

def fechar_modais(driver):
    """Remove todos os modais/iframes que possam estar bloqueando a tela"""
    try:
        driver.switch_to.default_content()
        driver.execute_script("""
            // Remove iframe_modal
            var iframe = document.getElementById('iframe_modal');
            if(iframe){ iframe.style.display='none'; iframe.src='about:blank'; }

            // Remove dialogModal
            var dm = document.getElementById('dialogModal');
            if(dm){ dm.style.display='none'; dm.classList.remove('in'); }

            // Remove todos os backdrops
            var bds = document.querySelectorAll('.modal-backdrop');
            bds.forEach(function(bd){ bd.parentNode.removeChild(bd); });

            // Remove modal-open do body
            document.body.classList.remove('modal-open');
            document.body.style.overflow = '';
            document.body.style.paddingRight = '';
        """)
        time.sleep(0.8)
    except: pass

def fechar_modal_inline(driver, dialog_id):
    """Fecha um modal inline do sistema antigo do Galileu (display:none)"""
    try:
        driver.execute_script(f"""
            var d = document.getElementById('{dialog_id}');
            if(d){{ d.style.display='none'; d.style.visibility='hidden'; }}
        """)
        time.sleep(0.3)
    except: pass

def fechar_modal_bootstrap(driver):
    """Fecha modais Bootstrap (backdrop, modal-dialog)"""
    try:
        # Tenta clicar no X ou botao Fechar do modal visivel
        for sel in [
            "//div[contains(@class,'modal') and contains(@style,'display: block')]//button[contains(@class,'close') or normalize-space()='×']",
            "//div[contains(@class,'modal') and contains(@style,'display: block')]//button[normalize-space()='Fechar']",
        ]:
            els = driver.find_elements(By.XPATH, sel)
            for el in els:
                try:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].click();", el)
                        time.sleep(0.5)
                        break
                except: continue
    except: pass
    fechar_modais(driver)

def aguardar_iframe_fechar(driver, tempo=8):
    """Aguarda o iframe_modal ficar oculto (display:none ou removido do DOM)"""
    try:
        WebDriverWait(driver, tempo).until(lambda d: d.execute_script("""
            var f = document.getElementById('iframe_modal');
            if(!f) return true;
            var s = window.getComputedStyle(f);
            return s.display === 'none' || s.visibility === 'hidden'
                   || !f.offsetParent || f.src === 'about:blank' || f.src === '';
        """))
    except: pass

def buscar_em_modal(driver, dialog_id, campo_id, localizar_id, termo, resultado_xpath):
    """
    Fluxo generico de busca em modal inline do Galileu:
    1. Digita no campo_id
    2. Clica no botao localizar_id
    3. Clica no primeiro resultado
    """
    campo = WV(driver, By.ID, campo_id, t=8)
    campo.clear(); campo.send_keys(termo); time.sleep(0.4)
    clicar(driver, driver.find_element(By.ID, localizar_id))
    time.sleep(1.5)
    res = driver.find_elements(By.XPATH, resultado_xpath)
    if res:
        clicar(driver, res[0]); time.sleep(1); return True
    return False

# ── CAMPOS ────────────────────────────────────────────────

def preencher_titulo(driver, titulo):
    try:
        # Aguarda form carregar
        WebDriverWait(driver, ESPERA).until(EC.presence_of_element_located((By.TAG_NAME,"form")))
        time.sleep(1.5)
        fechar_modais(driver)

        campo = None
        for by, sel in [
            (By.ID,    "titulo-biblioteca-tx_titulo"),
            (By.XPATH, "//input[contains(@id,'titulo') and contains(@id,'tx')]"),
            (By.XPATH, "//label[contains(.,'Título') or contains(.,'Titulo')]/following::input[@type='text'][1]"),
            (By.XPATH, "(//input[@type='text'])[1]"),
        ]:
            try: campo = WC(driver, by, sel, t=4); break
            except: continue

        if not campo: log("  Titulo: campo nao encontrado","ERRO"); return
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", campo)
        campo.clear()
        # Limpa o titulo removendo " PDF" e espacos extras
        titulo_limpo = titulo.replace(" PDF","").replace("  "," ").strip()
        campo.send_keys(titulo_limpo)
        log(f"  Titulo: {titulo_limpo[:55]}","OK")
    except Exception as e:
        log(f"  Erro Titulo: {e}","ERRO"); logging.exception("Titulo")

def sel_tipo_material(driver):
    try:
        sel = Select(W(driver, By.ID, "titulo-biblioteca-id_tipomaterialbiblioteca"))
        for o in sel.options:
            if "LIVRO" in o.text.upper():
                sel.select_by_visible_text(o.text)
                log(f"  Tipo Material: {o.text}","OK"); return
        log("  LIVRO nao encontrado","AVISO")
    except Exception as e:
        log(f"  Erro Tipo Material: {e}","ERRO"); logging.exception("TipoMaterial")

def sel_classificacao(driver):
    try:
        sel = Select(W(driver, By.XPATH,
            "//select[contains(@id,'classificacao') or contains(@id,'etaria')]"))
        opts = [o.text for o in sel.options]
        log(f"  Classificacao opcoes: {opts}")
        for o in sel.options:
            if "LIVRE" in o.text.upper():
                sel.select_by_visible_text(o.text)
                log(f"  Classificacao: {o.text}","OK"); return
        log(f"  Livre nao encontrado: {opts}","AVISO")
    except Exception as e:
        log(f"  Erro Classificacao: {e}","ERRO"); logging.exception("Classificacao")

def sel_editora(driver, editora):
    editora = editora.strip()
    if not editora: return
    try:
        sel = Select(W(driver, By.ID, "titulo-biblioteca-id_editora"))
        eu = editora.upper()
        for o in sel.options:
            if not o.text.strip(): continue
            if eu in o.text.upper() or o.text.upper() in eu:
                sel.select_by_visible_text(o.text)
                log(f"  Editora: {o.text}","OK"); return

        # Nao encontrou — cadastra via iframe
        log(f"  Cadastrando editora: {editora[:40]}","AVISO")
        driver.execute_script("openModalEditoraBiblioteca()")
        time.sleep(2.5)

        iframe = W(driver, By.ID, "iframe_modal", t=6)
        driver.switch_to.frame(iframe); time.sleep(1.5)

        campo = W(driver, By.XPATH, "//input[@type='text'][1]", t=5)
        campo.clear(); campo.send_keys(editora); time.sleep(0.4)

        btn_save = WC(driver, By.XPATH,
            "//button[normalize-space()='Salvar'] | //a[normalize-space()='Salvar']", t=5)
        clicar(driver, btn_save); time.sleep(2.5)

        # Tenta clicar Fechar dentro do iframe antes de sair
        try:
            btn_fechar = driver.find_element(By.XPATH,
                "//button[normalize-space()='Fechar'] | //button[@data-dismiss='modal']"
                " | //button[contains(@class,'close')]")
            driver.execute_script("arguments[0].click();", btn_fechar)
            time.sleep(0.5)
        except: pass

        driver.switch_to.default_content()
        time.sleep(1)

        # Força fechamento total do iframe e qualquer modal aberto
        driver.execute_script("""
            var iframe = document.getElementById('iframe_modal');
            if(iframe){
                iframe.style.display = 'none';
                iframe.style.visibility = 'hidden';
                iframe.style.zIndex = '-1';
                iframe.style.pointerEvents = 'none';
                iframe.src = 'about:blank';
            }
            var mw = document.getElementById('modal_win');
            if(mw){ mw.style.display='none'; mw.classList.remove('in'); }
            var dm = document.getElementById('dialogModal');
            if(dm){ dm.style.display='none'; dm.classList.remove('in'); }
            var backdrops = document.querySelectorAll('.modal-backdrop');
            backdrops.forEach(function(b){ b.parentNode.removeChild(b); });
            document.body.classList.remove('modal-open');
            document.body.style.overflow = '';
            document.body.style.paddingRight = '';
        """)
        time.sleep(3)
        log(f"  Editora cadastrada: {editora}","OK")

        # Re-seleciona
        sel2 = Select(W(driver, By.ID, "titulo-biblioteca-id_editora"))
        for o in sel2.options:
            if not o.text.strip(): continue
            if eu in o.text.upper() or o.text.upper() in eu:
                sel2.select_by_visible_text(o.text)
                log(f"  Editora selecionada: {o.text}","OK"); return
        sel2.select_by_index(len(sel2.options)-1)
        log("  Editora: ultima opcao","AVISO")
    except Exception as e:
        log(f"  Erro Editora: {e}","ERRO"); logging.exception("Editora")
        driver.switch_to.default_content()

def sel_idioma(driver):
    """
    Sempre seleciona PORTUGUÊS(BRASIL)
    O modal usa sistema antigo (display:none) — força exibicao via JS
    """
    try:
        fechar_modais(driver)
        aguardar_iframe_fechar(driver)
        time.sleep(1.5)

        driver.execute_script("window.scrollTo(0, 600);")
        time.sleep(0.5)

        # Clica no botao Selecionar de Idioma
        driver.execute_script("document.getElementById('btn-selecionar-idioma').click();")
        time.sleep(2)

        # Forca o modal ficar visivel (o sistema antigo usa display:none no dialog)
        driver.execute_script("""
            // Tenta os dois sistemas de modal (novo bootstrap e antigo div)
            var dialogs = [
                document.getElementById('dialog-selecionar-idioma'),
                document.querySelector('.modal-dialog[id*="idioma"]'),
                document.querySelector('[id*="idioma"][style*="display"]')
            ];
            dialogs.forEach(function(d){
                if(d){ d.style.display='block'; d.style.visibility='visible'; }
            });
            // Forca o campo input aparecer
            var campo = document.getElementById('selecionar-idioma-tx_idiomaBiblioteca');
            if(campo){ campo.style.display='block'; campo.style.visibility='visible'; }
        """)
        time.sleep(1)

        # Tenta encontrar o campo de busca de varias formas
        campo = None
        for by, sel in [
            (By.ID, "selecionar-idioma-tx_idiomaBiblioteca"),
            (By.XPATH, "//input[contains(@id,'selecionar-idioma')]"),
            (By.XPATH, "//input[contains(@name,'selecionar-idioma')]"),
            (By.XPATH, "//div[@id='dialog-selecionar-idioma']//input[@type='text']"),
            (By.XPATH, "//div[contains(@class,'modal')]//input[contains(@id,'idioma')]"),
        ]:
            try:
                el = driver.find_element(by, sel)
                if el:
                    campo = el; break
            except: continue

        if not campo:
            log("  Idioma: campo nao encontrado","ERRO"); return

        driver.execute_script("arguments[0].style.display='block';arguments[0].style.visibility='visible';", campo)
        fechar_modal_bootstrap(driver)
        time.sleep(0.2)
        # Seta o valor via JS em vez de click()+send_keys — esse campo eh
        # instavel (ElementClickIntercepted/NotInteractable intermitentes),
        # setar direto e disparar input/change eh equivalente pro Localizar.
        driver.execute_script("""
            var c = arguments[0];
            c.value = arguments[1];
            c.dispatchEvent(new Event('input', {bubbles:true}));
            c.dispatchEvent(new Event('change', {bubbles:true}));
            c.dispatchEvent(new Event('keyup', {bubbles:true}));
        """, campo, "Portugu")
        time.sleep(0.5)

        # Clica Localizar
        driver.execute_script("document.getElementById('btnLocalizarSelecionarIdioma').click();")
        time.sleep(2.5)

        # Inicializa selecionado antes das estratégias
        selecionado = False

        # Estratégia 1: acha a linha com BRASIL e clica no botão/label dela
        linhas = driver.find_elements(By.XPATH,
            "//div[@id='dialog-selecionar-idioma']//tr"
            " | //div[@id='grid-busca-selecionar-idioma']//tr"
        )
        for linha in linhas:
            try:
                if "BRASIL" in linha.text.upper():
                    btn_sel = linha.find_element(By.XPATH,
                        ".//label[contains(@class,'btnSelecionar')]"
                        " | .//button"
                        " | .//label[contains(@class,'btn')]"
                    )
                    driver.execute_script("arguments[0].click();", btn_sel)
                    time.sleep(1)
                    fechar_modal_inline(driver, "dialog-selecionar-idioma")
                    log("  Idioma: PORTUGUÊS(BRASIL)","OK")
                    selecionado = True; break
            except: continue

        # Estratégia 2: XPath direto pela célula com BRASIL
        if not selecionado:
            try:
                btn = driver.find_element(By.XPATH,
                    "//*[contains(text(),'BRASIL')]/ancestor::tr"
                    "/descendant::label[contains(@class,'btnSelecionar')]"
                    " | //*[contains(text(),'BRASIL')]/ancestor::tr"
                    "/descendant::button[contains(normalize-space(),'Selecionar')]"
                )
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
                fechar_modal_inline(driver, "dialog-selecionar-idioma")
                log("  Idioma: PORTUGUÊS(BRASIL) (estrategia 2)","OK")
                selecionado = True
            except: pass

        # Estratégia 3: pega o 4o label/botão Selecionar
        if not selecionado:
            todos = driver.find_elements(By.XPATH,
                "//div[@id='dialog-selecionar-idioma']//label[contains(@class,'btnSelecionar')]"
                " | //div[@id='dialog-selecionar-idioma']//button[contains(normalize-space(),'Selecionar')]"
            )
            log(f"  Idioma: {len(todos)} opcoes encontradas")
            if len(todos) >= 4:
                driver.execute_script("arguments[0].click();", todos[3])
                time.sleep(1)
                fechar_modal_inline(driver, "dialog-selecionar-idioma")
                log("  Idioma: PORTUGUÊS(BRASIL) (pos 4)","OK")
            elif todos:
                driver.execute_script("arguments[0].click();", todos[-1])
                time.sleep(1)
                fechar_modal_inline(driver, "dialog-selecionar-idioma")
                log(f"  Idioma: selecionou ultima opcao ({len(todos)})","OK")
            else:
                fechar_modal_inline(driver, "dialog-selecionar-idioma")
                log("  Idioma: nenhum resultado","AVISO")

    except Exception as e:
        log(f"  Erro Idioma: {type(e).__name__} — {str(e)[:150]}","ERRO")
        logging.exception("Idioma")
        try:
            driver.find_element(By.XPATH,"//button[normalize-space()='Fechar']").click()
        except: pass

def sel_autor(driver, autores_str):
    """
    Modal INLINE — id="dialog-selecionar-autor"
    Botao Selecionar: id="btn-selecionar-autor"
    Campo busca: id="selecionar-autor-tx_nomautorbiblioteca"
    Se nao achar: cria via iframe e depois busca novamente para selecionar
    """
    autor = autores_str.split(",")[0].strip()
    for s in ["(Organizador)","(Organizadora)","(Coordenador)","(Coordenadora)","et al","PDF"]:
        autor = autor.replace(s,"").strip()
    if not autor or len(autor)<3:
        log("  Autor: nome invalido","AVISO"); return

    # Usa as 2 primeiras palavras para busca (mais preciso que só a primeira)
    palavras = autor.split()
    if len(palavras) >= 2:
        termo = palavras[0] + " " + palavras[1]
    else:
        termo = palavras[0]

    def buscar_e_selecionar(termo_busca=None):
        """Abre o modal de busca, pesquisa e clica em Selecionar. Retorna True se achou."""
        t = termo_busca or termo
        try:
            fechar_modais(driver)
            time.sleep(1.5)

            driver.execute_script(
                "document.getElementById('btn-selecionar-autor').click();"
            )
            time.sleep(2)

            driver.execute_script("""
                var d = document.getElementById('dialog-selecionar-autor');
                if(d){ d.style.display='block'; d.style.visibility='visible'; d.style.zIndex='9999'; }
                var c = document.getElementById('selecionar-autor-tx_nomautorbiblioteca');
                if(c){ c.style.display='block'; c.style.visibility='visible'; c.style.zIndex='9999'; }
            """)
            time.sleep(0.8)

            # Limpa o campo completamente antes de digitar
            driver.execute_script("""
                var c = document.getElementById('selecionar-autor-tx_nomautorbiblioteca');
                if(c){ c.value = ''; c.focus(); }
            """)
            time.sleep(0.3)

            # Remove acentos do termo para compatibilidade com o sistema
            import unicodedata
            termo_ascii = ''.join(
                c for c in unicodedata.normalize('NFD', t)
                if unicodedata.category(c) != 'Mn'
            )
            # Usa só as primeiras 8 letras sem acento para busca mais ampla
            termo_busca_js = termo_ascii[:8]

            # Digita o termo via send_keys (mais confiável que JS value)
            campo = driver.find_element(By.ID, "selecionar-autor-tx_nomautorbiblioteca")
            # Limpa com JS primeiro, depois com clear(), depois digita
            driver.execute_script("arguments[0].value = '';", campo)
            campo.clear()
            campo.send_keys(termo_busca_js)
            time.sleep(0.5)

            # Clica Localizar
            driver.execute_script("""
                var btns = document.querySelectorAll('#dialog-selecionar-autor button');
                for(var i=0;i<btns.length;i++){
                    if(btns[i].textContent.trim().indexOf('Localizar')>=0){
                        btns[i].click(); break;
                    }
                }
            """)
            time.sleep(2)

            res = driver.find_elements(By.XPATH,
                "//div[@id='dialog-selecionar-autor']//label[contains(@class,'btnSelecionar')]"
                " | //div[@id='dialog-selecionar-autor']//button[contains(normalize-space(),'Selecionar')]"
                " | //div[@id='dialog-selecionar-autor']//*[contains(@class,'btnSelecionar')]"
            )
            log(f"  Autor busca '{t}': {len(res)} resultado(s)")
            if res:
                driver.execute_script("arguments[0].click();", res[0])
                time.sleep(1)
                fechar_modal_inline(driver, "dialog-selecionar-autor")
                return True

            fechar_modal_inline(driver, "dialog-selecionar-autor")
            return False
        except Exception as e:
            log(f"  Erro buscar_selecionar: {type(e).__name__}: {str(e)[:120]}","AVISO")
            logging.exception("buscar_e_selecionar")
            fechar_modal_inline(driver, "dialog-selecionar-autor")
            fechar_modais(driver)
            return False

    # Tenta buscar com varios termos antes de decidir criar
    palavras = autor.split()
    termos_para_tentar = []

    # Termo 1: 2 primeiras palavras
    if len(palavras) >= 2:
        termos_para_tentar.append(palavras[0] + " " + palavras[1])
    # Termo 2: só a primeira palavra
    termos_para_tentar.append(palavras[0])
    # Termo 3: primeiras 3 palavras se disponível
    if len(palavras) >= 3:
        termos_para_tentar.insert(1, palavras[0] + " " + palavras[1] + " " + palavras[2])

    # Remove duplicatas mantendo ordem
    vistos = set()
    termos_unicos = []
    for t in termos_para_tentar:
        if t not in vistos:
            vistos.add(t); termos_unicos.append(t)

    try:
        # Tenta buscar com todos os termos antes de criar
        for t in termos_unicos:
            log(f"  Buscando autor com termo: '{t}'")
            if buscar_e_selecionar(t):
                log(f"  Autor: {autor[:50]}","OK"); return

        # Definitivamente nao existe — cria novo
        log(f"  Autor nao encontrado em {termos_unicos}, criando...","AVISO")
        fechar_modais(driver)

        btn_inc = WC(driver, By.XPATH,
            "//button[contains(@onclick,'openModalAutorBiblioteca')]", t=5)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_inc)
        clicar(driver, btn_inc); time.sleep(3)

        # Entra no iframe
        iframe = W(driver, By.ID, "iframe_modal", t=8)
        driver.switch_to.frame(iframe)
        time.sleep(3)

        # Fecha popup "Atenção! Não se esqueça de clicar em SALVAR"
        for _ in range(3):
            try:
                btn_ok = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH,
                        "//*[normalize-space()='Ok' or normalize-space()='OK' or normalize-space()='ok']"
                        "[self::button or self::a or self::input]"
                    ))
                )
                driver.execute_script("arguments[0].click();", btn_ok)
                log("  Popup Ok fechado","OK")
                time.sleep(1)
                break
            except: break

        # Aguarda campo aparecer após fechar popup
        time.sleep(1)
        campo_nome = None
        for by, sel in [
            (By.ID,    "autor-biblioteca-tx_nomeautorbiblioteca"),
            (By.XPATH, "//input[contains(@id,'nomeautor') or contains(@name,'nomeautor')]"),
            (By.XPATH, "//input[contains(@id,'nome') or contains(@name,'nome')]"),
            (By.XPATH, "//input[@type='text'][not(ancestor::*[contains(@style,'display:none') or contains(@style,'display: none')])]"),
            (By.XPATH, "//input[@type='text'][1]"),
        ]:
            try:
                el = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((by, sel)))
                if el:
                    campo_nome = el
                    log(f"  Campo autor encontrado: {sel[:50]}")
                    break
            except: continue

        if not campo_nome:
            log("  Campo nome autor nao encontrado no iframe","AVISO")
            driver.switch_to.default_content()
            fechar_modais(driver); return

        campo_nome.clear(); campo_nome.send_keys(autor); time.sleep(0.4)
        btn_save = WC(driver, By.XPATH,
            "//button[normalize-space()='Salvar'] | //a[normalize-space()='Salvar']"
            " | //input[@value='Salvar']", t=5)
        clicar(driver, btn_save)
        time.sleep(2)

        # Tenta clicar em Fechar dentro do iframe (caso apareça após salvar)
        try:
            btn_fechar = driver.find_element(By.XPATH,
                "//button[normalize-space()='Fechar'] | //a[normalize-space()='Fechar']"
                " | //button[contains(@class,'close')] | //button[@data-dismiss='modal']")
            clicar(driver, btn_fechar)
            time.sleep(1)
        except: pass

        # Volta ao contexto principal
        driver.switch_to.default_content()
        time.sleep(2)

        # Fecha popup "Atenção! Não se esqueça de clicar em SALVAR" se apareceu
        try:
            btn_ok = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//button[normalize-space()='Ok' or normalize-space()='OK' or normalize-space()='ok']"
                ))
            )
            driver.execute_script("arguments[0].click();", btn_ok)
            log("  Popup Ok fechado","OK")
            time.sleep(1)
        except: pass

        # Força fechamento completo do iframe e modal_win via JS
        driver.execute_script("""
            var iframe = document.getElementById('iframe_modal');
            if(iframe){
                iframe.style.display = 'none';
                iframe.style.visibility = 'hidden';
                iframe.style.zIndex = '-1';
                iframe.style.pointerEvents = 'none';
                iframe.src = 'about:blank';
            }
            var mw = document.getElementById('modal_win');
            if(mw){ mw.style.display='none'; mw.classList.remove('in'); }
            var backdrops = document.querySelectorAll('.modal-backdrop');
            backdrops.forEach(function(b){ b.parentNode.removeChild(b); });
            document.body.classList.remove('modal-open');
            document.body.style.overflow = '';
            document.body.style.paddingRight = '';
        """)
        time.sleep(2)
        log(f"  Autor criado: {autor[:50]}","OK")

        # Busca e seleciona o autor recem-criado
        if buscar_e_selecionar():
            log(f"  Autor selecionado: {autor[:50]}","OK")
        else:
            log(f"  Autor criado mas nao selecionado","AVISO")
            # Ultima tentativa: recarrega a pagina nao — so loga o aviso
            log(f"  ATENCAO: autor '{autor[:40]}' foi criado mas NAO selecionado — livro sera recusado","ERRO")

    except Exception as e:
        log(f"  Erro Autor: {e}","ERRO"); logging.exception("Autor")
        driver.switch_to.default_content()
        fechar_modais(driver)

def sel_genero(driver, genero):
    """
    Modal INLINE — id="dialog-selecionar-genero-literario"
    Botao: id="btn-selecionar-genero-literario"
    Campo: id="selecionar-genero-literario-tx_generoliteriobiblioteca"
    """
    mapa = {
        "TECNICO / CIENTIFICO (SAUDE)": "Técnico-Científico",
        "TECNICO/CIENTIFICO":           "Técnico-Científico",
        "TECNICO-CIENTIFICO":           "Técnico-Científico",
        "ACADEMICO":  "acadêmico",
        "TECNICO":    "Técnico",
        "DESCRITIVO": "Descritivo",
    }
    g = genero.upper().replace("Ê","E").replace("Â","A").replace("É","E").replace("Í","I")
    busca = mapa.get(g, genero.split("/")[0].strip())

    try:
        fechar_modais(driver)
        time.sleep(1)

        driver.execute_script("window.scrollTo(0, 800);")
        time.sleep(0.5)

        btn_sel = WC(driver, By.ID, "btn-selecionar-genero-literario", t=10)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_sel)
        time.sleep(0.5)
        clicar(driver, btn_sel)
        time.sleep(2)

        # Forca o modal ficar visivel (mesmo sistema antigo do idioma)
        driver.execute_script("""
            var dialog = document.getElementById('dialog-selecionar-genero-literario');
            if(dialog){ dialog.style.display='block'; dialog.style.visibility='visible'; }
            var campo = document.getElementById('selecionar-genero-literario-tx_generoliteriobiblioteca');
            if(campo){ campo.style.display='block'; campo.style.visibility='visible'; }
        """)
        time.sleep(0.5)

        # Tenta encontrar campo de busca
        campo = None
        for by, sel in [
            (By.ID, "selecionar-genero-literario-tx_generoliteriobiblioteca"),
            (By.XPATH, "//input[contains(@id,'selecionar-genero')]"),
            (By.XPATH, "//div[@id='dialog-selecionar-genero-literario']//input[@type='text']"),
        ]:
            try:
                el = driver.find_element(by, sel)
                if el: campo = el; break
            except: continue

        if not campo:
            log("  Genero: campo nao encontrado","AVISO"); return

        driver.execute_script("arguments[0].style.display='block';arguments[0].style.visibility='visible';", campo)

        for termo in [busca[:6], "Tecni", "acad"]:
            campo.clear(); campo.send_keys(termo); time.sleep(0.4)

            try:
                btn_loc = driver.find_element(By.XPATH,
                    "//div[@id='dialog-selecionar-genero-literario']"
                    "//button[contains(normalize-space(),'Localizar')]"
                    " | //div[@id='dialog-selecionar-genero-literario']"
                    "//input[@type='button']"
                )
                clicar(driver, btn_loc)
            except:
                driver.execute_script("document.getElementById('btnLocalizarSelecionarGeneroLiterario') && document.getElementById('btnLocalizarSelecionarGeneroLiterario').click();")
            time.sleep(1.5)

            res = driver.find_elements(By.XPATH,
                "//div[@id='dialog-selecionar-genero-literario']//label[contains(@class,'btnSelecionar')]"
                " | //div[@id='dialog-selecionar-genero-literario']//button[contains(normalize-space(),'Selecionar')]"
                " | //div[@id='dialog-selecionar-genero-literario']//*[contains(@class,'btnSelecionar')]"
            )
            if res:
                clicar(driver, res[0]); time.sleep(1)
                fechar_modal_inline(driver, "dialog-selecionar-genero-literario")
                log(f"  Genero: {busca}","OK"); return

        fechar_modal_inline(driver, "dialog-selecionar-genero-literario")
        log("  Genero: nenhum resultado encontrado","AVISO")
    except Exception as e:
        log(f"  Erro Genero: {e}","ERRO"); logging.exception("Genero")

def fazer_upload(driver, nome_arquivo):
    """
    Upload do PDF — input id="arquivo" (oculto, revelado via JS)
    """
    if not nome_arquivo:
        log("  Arquivo: nao informado na planilha","AVISO"); return False

    caminho = os.path.join(PASTA_LIVROS, nome_arquivo)
    if not os.path.exists(caminho):
        log(f"  Arquivo nao encontrado: {nome_arquivo}","AVISO"); return False

    try:
        fechar_modais(driver)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)

        inp = driver.find_element(By.ID, "arquivo")
        # Revela o input file oculto
        driver.execute_script(
            "arguments[0].style.opacity='1';"
            "arguments[0].style.display='block';"
            "arguments[0].style.position='relative';"
            "arguments[0].style.zIndex='9999';"
            "arguments[0].style.left='0';"
            "arguments[0].style.top='0';", inp)
        inp.send_keys(caminho)
        time.sleep(2.5)
        log(f"  Upload PDF: {nome_arquivo}","OK")
        return True
    except Exception as e:
        log(f"  Erro Upload: {e}","ERRO"); logging.exception("Upload"); return False

def salvar(driver):
    try:
        fechar_modais(driver)
        driver.execute_script("window.scrollTo(0,0);")
        time.sleep(0.8)

        btn = WC(driver, By.ID, "btnSalvar", t=8)
        clicar(driver, btn)
        time.sleep(8)  # Aguarda mais tempo para uploads grandes (até 20MB)

        # Verifica se apareceu mensagem de erro REAL na pagina
        # Ignora textos informativos que nao sao erros
        TEXTOS_IGNORAR = [
            "caso não informado",
            "prazo limite para devolução",
            "configurações",
            "×",  # botao fechar do modal de sucesso
            "x",
        ]
        erros = driver.find_elements(By.XPATH,
            "//*[contains(@class,'alert-danger') or contains(@class,'msg-erro')]"
            " | //*[contains(@id,'containerMessage') and (contains(.,'erro') or contains(.,'obrigat') or contains(.,'inválid'))]"
        )
        for e in erros:
            try:
                if not e.is_displayed(): continue
                txt = e.text.strip().lower()
                if not txt: continue
                # Ignora textos informativos
                if any(ig in txt for ig in TEXTOS_IGNORAR): continue
                log(f"  Erro ao salvar: {e.text.strip()[:100]}","ERRO")
                return False
            except: continue

        # Verifica se salvou com sucesso
        url_atual = driver.current_url

        # Se redirecionou para listagem = sucesso
        if "formulario" not in url_atual and "titulo-biblioteca" in url_atual:
            log("  Salvo! (redirecionou para lista)","OK")
            return True

        # Verifica mensagem de sucesso
        try:
            msg = driver.find_element(By.XPATH,
                "//*[contains(@class,'alert-success') or contains(@class,'msg-sucesso')]"
                " | //*[contains(@id,'containerMessage') and (contains(.,'sucesso') or contains(.,'salvo') or contains(.,'cadastr'))]"
            )
            if msg.is_displayed() and msg.text.strip() not in ("×","x",""):
                log(f"  Salvo com confirmacao: {msg.text.strip()[:60]}","OK")
                return True
        except: pass

        # Verifica se o formulario agora tem um ID na URL (foi salvo e voltou ao form com ID)
        if "id=" in url_atual or "/editar/" in url_atual:
            log("  Salvo! (formulario com ID)","OK")
            return True

        # Nenhum sinal positivo de sucesso (nem redirecionamento, nem mensagem,
        # nem ID na URL) e nenhum erro visivel tambem. Uma auditoria manual
        # confirmou que ESSE caminho ambiguo quase sempre significa que o
        # clique em Salvar nao surtiu efeito (o "livro" nunca existiu de
        # verdade no Galileu, mesmo sem erro aparente) — entao trata como
        # falha em vez de assumir sucesso as cegas.
        log("  Sem confirmacao de sucesso nem erro visivel — tratando como falha","ERRO")
        return False

    except Exception as e:
        log(f"  Erro Salvar: {e}","ERRO")
        logging.exception("Salvar")
        return False

# ── CADASTRO PRINCIPAL ────────────────────────────────────

def cadastrar_livro(driver, livro, n, total):
    log(f"\n{'='*55}")
    log(f"[{n}/{total}] {livro['titulo'][:55]}","INICIO")
    log(f"{'='*55}")
    try:
        driver.get(URL_CADASTRO); time.sleep(4)
        if "login" in driver.current_url.lower():
            log("Sessao expirada! Faca login e ENTER...","AVISO"); input()

        fechar_modais(driver)

        preencher_titulo(driver, livro["titulo"]);      time.sleep(0.3)
        sel_tipo_material(driver);                       time.sleep(0.3)
        sel_classificacao(driver);                       time.sleep(0.3)
        sel_editora(driver, livro["editora"]);           time.sleep(0.5)

        driver.execute_script("window.scrollTo(0,700);"); time.sleep(1)
        fechar_modais(driver)

        sel_idioma(driver);                              time.sleep(0.8)
        sel_autor(driver, livro["autores"]);             time.sleep(0.8)
        sel_genero(driver, livro["genero"]);             time.sleep(0.8)
        fazer_upload(driver, livro["arquivo"]);          time.sleep(1)

        ok = salvar(driver)
        if ok:
            salvar_ok(livro["titulo"])
            log("CADASTRADO COM SUCESSO!","OK")
        else:
            log("Possivel erro ao salvar","AVISO")
        return ok
    except Exception as e:
        log(f"Erro geral: {e}","ERRO")
        logging.exception(f"Geral: {livro['titulo']}")
        driver.switch_to.default_content()
        return False

# ── MAIN ─────────────────────────────────────────────────

def main():
    print("\n"+"="*55)
    print("  CADASTRO AUTOMATICO - SISTEMA GALILEU")
    print("="*55)
    print(f"\n  Log de erros: {LOG_ERROS}\n")
    input("Chrome aberto e logado? Pressione ENTER...")

    livros = carregar_planilha()
    ja_ok  = carregar_cadastrados()
    novos  = [l for l in livros if l["titulo"] not in ja_ok]

    if not novos:
        log("Todos ja cadastrados!","OK"); return

    log(f"{len(novos)} novos | {len(ja_ok)} ja feitos")
    driver = iniciar_chrome()
    ok = falha = 0

    try:
        for i, livro in enumerate(novos, 1):
            if cadastrar_livro(driver, livro, i, len(novos)): ok += 1
            else: falha += 1
            time.sleep(2)
    except KeyboardInterrupt:
        log("Interrompido.","AVISO")
    finally:
        print("\n"+"="*55)
        log(f"Sucesso : {ok}","OK")
        if falha: log(f"Com erro: {falha}","AVISO")
        log(f"Log salvo em: {LOG_ERROS}")
        input("\nENTER para encerrar...")

if __name__ == "__main__":
    main()
