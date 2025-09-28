from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchWindowException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import threading
import queue
import random
import time
import os
import re
import sys
import argparse
import datetime
import tempfile
import atexit
import shutil


driver = None
wait = None
CURRENT_EMAIL = ""

# === Manejo de cancelación (Ctrl+Z / Ctrl+C) ===
import builtins as _builtins
_original_input = _builtins.input

def safe_input(prompt: str = ""):
    """Wrapper de input que permite cancelar con Ctrl+Z (EOF) o Ctrl+C.
    Ctrl+Z (EOF) en Windows -> EOFError
    Ctrl+C -> KeyboardInterrupt
    Ambos terminan el programa limpiamente.
    """
    try:
        return _original_input(prompt)
    except EOFError:
        print("\nCancelado por usuario (Ctrl+Z). Saliendo...")
        try:
            # Intentar cerrar drivers si existen
            if 'driver' in globals() and driver:
                driver.quit()
        except Exception:
            pass
        sys.exit(0)
    except KeyboardInterrupt:
        print("\nCancelado por usuario (Ctrl+C). Saliendo...")
        try:
            if 'driver' in globals() and driver:
                driver.quit()
        except Exception:
            pass
        sys.exit(0)

# Reemplazar input globalmente para todo el script
input = safe_input

# ========================= NUEVAS VARIABLES GLOBALES (inicialización diferida) =========================
FORM_URL = None
NUM_THREADS = None
MAX_SUBMISSIONS = None
EMAILS_FILE = None
used_path = None
email_queue = None
success_lock = threading.Lock()
success_count = 0
emails_to_use = []

# ========================= UTILIDADES PARA GENERAR CORREOS =========================
import unicodedata

def _strip_accents(s: str) -> str:
    try:
        return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    except Exception:
        return s

def build_email(nombre: str, apellido: str, sufijo: str, dominio: str) -> str:
    base = f"{nombre}{apellido}{sufijo}".lower().replace(' ', '')
    base = _strip_accents(base)
    dominio = dominio.lower().strip()
    if dominio.startswith('@'):
        dominio = dominio[1:]
    return f"{base}@{dominio}"

def generate_emails_interactive():
    print("\n=== GENERADOR DE CORREOS ===")
    print("Puedes escribir valores separados por coma o la ruta a un archivo .txt (uno por línea). Ej: nombres.txt")

    def load_list(label: str, obligatorio=True, allow_empty=False, validate=None):
        while True:
            raw = input(label).strip()
            if not raw:
                if not obligatorio:
                    return []
                print("Este campo es obligatorio.")
                continue
            # ¿Es archivo?
            if os.path.isfile(raw) and raw.lower().endswith('.txt'):
                try:
                    with open(raw, 'r', encoding='utf-8') as f:
                        items = [l.strip() for l in f if l.strip()]
                except Exception as e:
                    print(f"No se pudo leer el archivo: {e}")
                    continue
                if not items and not allow_empty:
                    print("El archivo está vacío.")
                    continue
                if validate:
                    filtered = []
                    for it in items:
                        if validate(it):
                            filtered.append(it)
                        else:
                            print(f"Aviso: '{it}' ignorado por formato inválido")
                    items = filtered
                if items or allow_empty:
                    return items
                print("No hay elementos válidos en el archivo.")
                continue
            # No es archivo: separar por coma / punto y coma
            parts = [p.strip() for p in re.split(r'[;,]', raw) if p.strip()]
            if not parts:
                if allow_empty:
                    return []
                print("Ingresa al menos un valor.")
                continue
            if validate:
                filtered = []
                for it in parts:
                    if validate(it):
                        filtered.append(it)
                    else:
                        print(f"Aviso: '{it}' ignorado por formato inválido")
                parts = filtered
            if parts or allow_empty:
                return parts

    def maybe_pick_file(label_base, obligatorio=True, allow_empty=False, validate=None):
        choice = input(f"¿Quieres abrir un archivo .txt para {label_base.lower()}? (s/N): ").strip().lower()
        if choice in ("s", "si", "sí", "y", "yes"):
            path = _open_file_dialog(f"Selecciona archivo de {label_base}")
            if path and os.path.isfile(path):
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        data = [l.strip() for l in f if l.strip()]
                except Exception as e:
                    print(f"Error leyendo archivo: {e}. Se pedirá manualmente.")
                    data = []
                if validate:
                    filtered = []
                    for it in data:
                        if validate(it):
                            filtered.append(it)
                        else:
                            print(f"Aviso: '{it}' ignorado por formato inválido")
                    data = filtered
                if data or allow_empty:
                    print(f"Cargados {len(data)} {label_base.lower()} desde archivo.")
                    return data
                else:
                    print("El archivo no tenía elementos válidos; ingrésalos manualmente.")
        return load_list(f"{label_base}: ", obligatorio=obligatorio, allow_empty=allow_empty, validate=validate)

    nombres_apellidos_pares = []  # lista de tuplas (nombres_compuestos, apellidos_compuestos)
    nombres = []
    apellidos = []

    # Primero capturamos nombres; si el usuario elige archivo, ofrecer detectar formato de 4 partes
    posibles_nombres = maybe_pick_file("Nombres", obligatorio=True)
    detect_full = False
    if posibles_nombres and all(len(line.split()) >= 4 for line in posibles_nombres):
        # Preguntar si el usuario quiso formato completo
        resp_full = input("Se detectaron líneas con >= 4 partes. ¿Interpretar como 'Nombre1 Nombre2 Apellido1 Apellido2'? (s/N): ").strip().lower()
        if resp_full in ("s","si","sí","y","yes"):
            detect_full = True
    if detect_full:
        def _clean_part(p: str) -> str:
            p = p.strip()
            # Mantener letras y caracteres acentuados / ñ
            p = re.sub(r"[^A-Za-zÁÉÍÓÚÜáéíóúüÑñ]", "", p)
            if not p:
                return p
            return p[0].upper() + p[1:].lower()
        skipped = 0
        for line in posibles_nombres:
            raw = line.strip()
            if not raw:
                continue
            tokens = raw.split()
            if len(tokens) < 4:
                skipped += 1
                continue
            # Si hay más de 4 tokens, asumir primeros 2 nombres y últimos 2 apellidos
            if len(tokens) > 4:
                tokens = tokens[:2] + tokens[-2:]
            cleaned = [_clean_part(t) for t in tokens]
            if any(not c for c in cleaned):
                skipped += 1
                continue
            nom_comp = " ".join(cleaned[:2])
            ape_comp = " ".join(cleaned[2:4])
            nombres_apellidos_pares.append((nom_comp, ape_comp))
        print(f"Detectados {len(nombres_apellidos_pares)} nombres completos (2 nombres + 2 apellidos). Se omitirá la petición de apellidos por separado. Omitidos: {skipped}")
    else:
        nombres = posibles_nombres
        apellidos = maybe_pick_file("Apellidos", obligatorio=True)
    years = maybe_pick_file(
        "Años (opcional, 2-4 dígitos) o Enter para aleatorios",
        obligatorio=False,
        allow_empty=True,
        validate=lambda x: bool(re.match(r'^\d{2,4}$', x))
    )
    use_random_numbers = not years

    if use_random_numbers:
        try:
            per_combo = input("¿Cuántos correos generar por combinación nombre+apellido? (default 3): ").strip()
            per_combo = int(per_combo) if per_combo else 3
            per_combo = max(1, min(50, per_combo))
        except Exception:
            per_combo = 3
    else:
        per_combo = 1  # cada combinación año ya produce uno por defecto

    dominios = maybe_pick_file("Dominio(s)", obligatorio=True)

    choose_save = input("¿Deseas elegir ruta y nombre del archivo de salida con un diálogo? (s/N): ").strip().lower()
    if choose_save in ("s", "si", "sí", "y", "yes"):
        # Reusar diálogo de abrir adaptado a guardar si disponible
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk(); root.withdraw()
            try:
                root.attributes("-topmost", True)
            except Exception:
                pass
            path = filedialog.asksaveasfilename(title="Guardar lista de correos", initialfile="emails_generados.txt", defaultextension=".txt", filetypes=[("Archivos de texto", "*.txt"), ("Todos los archivos", "*.*")])
            root.destroy()
            if path:
                out_path = os.path.abspath(path)
            else:
                print("No se eligió archivo, se usará emails_generados.txt en el directorio actual.")
                out_path = os.path.abspath("emails_generados.txt")
        except Exception:
            out_path = os.path.abspath("emails_generados.txt")
    else:
        out_path = input("Nombre del archivo de salida (Enter para emails_generados.txt): ").strip() or "emails_generados.txt"
        out_path = os.path.abspath(out_path)

    generated = set()
    if nombres_apellidos_pares:
        for nombre_comp, apellido_comp in nombres_apellidos_pares:
            if years:
                for y in years:
                    for dom in dominios:
                        email = build_email(nombre_comp.replace(' ', ''), apellido_comp.replace(' ', ''), y, dom)
                        generated.add(email)
            else:
                for dom in dominios:
                    count = 0
                    tries = 0
                    while count < per_combo and tries < per_combo * 10:
                        tries += 1
                        length = random.choice([2, 3, 4])
                        suf = ''.join(random.choice('0123456789') for _ in range(length))
                        email = build_email(nombre_comp.replace(' ', ''), apellido_comp.replace(' ', ''), suf, dom)
                        if email not in generated:
                            generated.add(email)
                            count += 1
    else:
        for nombre in nombres:
            for apellido in apellidos:
                if years:
                    for y in years:
                        for dom in dominios:
                            email = build_email(nombre, apellido, y, dom)
                            generated.add(email)
                else:
                    for dom in dominios:
                        count = 0
                        tries = 0
                        while count < per_combo and tries < per_combo * 10:
                            tries += 1
                            length = random.choice([2, 3, 4])
                            suf = ''.join(random.choice('0123456789') for _ in range(length))
                            email = build_email(nombre, apellido, suf, dom)
                            if email not in generated:
                                generated.add(email)
                                count += 1

    emails_list = sorted(generated)
    with open(out_path, 'w', encoding='utf-8') as f:
        for e in emails_list:
            f.write(e + '\n')

    print(f"Se generaron {len(emails_list)} correos y se guardaron en: {out_path}")
    return out_path

def menu_principal():
    print("""
==============================
  BOT FORMULARIO - MENÚ
==============================
1) Usar bot con lista de correos existente
2) Crear lista de correos (generador) y luego usar bot
==============================
""")
    op = input("Selecciona una opción (1/2): ").strip()
    while op not in ('1','2'):
        op = input("Opción inválida. Selecciona 1 o 2: ").strip()
    return op

def initialize_emails():
    """Inicializa la cola y variables relacionadas a correos después de definir EMAILS_FILE."""
    global used_path, emails_to_use, email_queue, success_count
    if not EMAILS_FILE:
        raise RuntimeError("EMAILS_FILE no está definido para inicializar correos.")
    used_path = os.path.splitext(EMAILS_FILE)[0] + "_used.txt"
    emails_to_use = load_emails()
    if not emails_to_use:
        print("No hay correos disponibles (todos usados o archivo vacío). Genera más correos.")
        sys.exit(1)
    print(f"Correos disponibles: {len(emails_to_use)}")
    email_queue = queue.Queue()
    success_count = 0
    for email in emails_to_use:
        email_queue.put(email)


def _open_file_dialog(title):
    """Abre un cuadro de diálogo del sistema para seleccionar un archivo y devuelve la ruta elegida."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    try:
        selected = filedialog.askopenfilename(title=title, filetypes=[("Archivos de texto", "*.txt"), ("Todos los archivos", "*.*")])
    finally:
        root.destroy()
    return selected or None


def _force_fill_email_js(driver, email):
    """Intenta aplicar el valor del correo mediante JavaScript como último recurso."""
    script = r"""
        const value = arguments[0];
        const candidates = Array.from(document.querySelectorAll('input, textarea'));
        for (const el of candidates) {
            if (!el || el.disabled || el.readOnly) continue;
            const text = [
                el.getAttribute('aria-label'),
                el.getAttribute('name'),
                el.getAttribute('placeholder'),
                el.getAttribute('autocomplete')
            ].filter(Boolean).join(' ').toLowerCase();
            if (text.includes('correo') || text.includes('email') || text.includes('mail') || el.type === 'email') {
                try {
                    el.focus();
                } catch (err) {
                    /* ignore */
                }
                el.value = value;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
            }
        }
        return false;
    """
    try:
        return bool(driver.execute_script(script, email))
    except Exception:
        return False

# Pausas eliminadas para máxima velocidad
def human_pause(a=0, b=0):
    """Mantiene compatibilidad con llamadas de pausa sin introducir retardos reales."""
    pass  # Sin pausas

# ===== Configuración inicial =====
MAX_SUBMISSIONS = None  # Limite opcional de envíos; None para ilimitado
REUSE_SINGLE_BROWSER = True  # Reutiliza la misma ventana de Chrome para todos los envíos
ALLOW_MANUAL_LOGIN = True    # Si el formulario exige login, esperar a que el usuario inicie sesión manualmente

# Opcional: reutilizar un perfil de Chrome ya logueado para evitar prompt de inicio de sesión
# Configura variables de entorno:


def get_user_input(default_emails_file: str = None):
    """Obtiene configuración del usuario (CLI/env/interactive)"""
    print("=== CONFIGURACIÓN DEL BOT ===")

    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--url", dest="url")
    parser.add_argument("--threads", dest="threads")
    parser.add_argument("--max-submissions", dest="max_submissions", type=int)
    parser.add_argument("--emails", dest="emails")
    try:
        args, _ = parser.parse_known_args()
    except SystemExit:
        args = argparse.Namespace(url=None, threads=None, max_submissions=None)

    url = (args.url or os.environ.get("FORM_URL") or "").strip()
    threads_env = os.environ.get("NUM_THREADS")
    threads_arg = args.threads
    max_submissions_arg = args.max_submissions
    emails_arg = getattr(args, "emails", None)
    emails_env = os.environ.get("EMAILS_FILE")
    emails_file = (emails_arg or emails_env or default_emails_file or "").strip()

    non_interactive = not sys.stdin.isatty()

    # Solicitar URL obligatoriamente
    if not url:
        if non_interactive:
            print("ERROR: Debes proporcionar una URL del formulario.")
            print("Sugerencia: python bot.py --url=https://docs.google.com/forms/TU_FORMULARIO")
            print("También puedes definir la variable de entorno FORM_URL con la dirección del formulario")
            sys.exit(1)
        else:
            while not url:
                url = input("\nIngresa el enlace del formulario de Google (obligatorio): ").strip()
                if not url:
                    print("La URL es obligatoria. No puedes continuar sin proporcionar una dirección válida.")
                    print("Ejemplo: https://docs.google.com/forms/d/e/1FAIpQLSd.../viewform")

    # Validar formato de URL
    if not re.match(r'^https?://', url, re.IGNORECASE):
        print(f"URL inválida: '{url}'. Debe comenzar con http:// o https://")
        if non_interactive:
            sys.exit(1)
        else:
            print("Por favor, ingresa una URL válida.")
            return get_user_input()  # Reintentar

    # Normalizar ruta del archivo de correos si ya se proporcionó
    def _normalize_path(p: str) -> str:
        return os.path.abspath(os.path.expanduser(p))

    if emails_file:
        emails_file = _normalize_path(emails_file)

    # Solicitar archivo de correos
    while True:
        if emails_file and os.path.isfile(emails_file):
            break

        if emails_file:
            message = f"No se encontró el archivo de correos: {emails_file}"
        else:
            message = "Debes proporcionar un archivo con la lista de correos."

        if non_interactive:
            print(message)
            print("Sugerencia: python bot.py --emails=RUTA_DEL_ARCHIVO.txt")
            print("También puedes definir la variable de entorno EMAILS_FILE")
            sys.exit(1)

        print(message)
        dialog_choice = _open_file_dialog("Selecciona el archivo con la lista de correos")
        if dialog_choice:
            emails_file = _normalize_path(dialog_choice)
            continue
        emails_file = input("\nRuta del archivo con correos (uno por línea, obligatorio): ").strip().strip('"')
        if not emails_file:
            continue
        emails_file = _normalize_path(emails_file)

    print(f"Archivo de correos seleccionado: {emails_file}")

    def parse_threads(val):
        try:
            n = int(val)
            return max(1, min(10, n))
        except Exception:
            return None

    num_threads = parse_threads(threads_arg) or parse_threads(threads_env)
    if num_threads is None:
        if non_interactive:
            num_threads = 1
        else:
            while True:
                try:
                    raw = input("\n¿Cuántos formularios quieres llenar simultáneamente? (1-10, Enter para 1): ").strip()
                    if not raw:
                        num_threads = 1
                        break
                    num_threads = int(raw)
                    if 1 <= num_threads <= 10:
                        break
                    else:
                        print("Por favor ingresa un número entre 1 y 10")
                except ValueError:
                    print("Por favor ingresa un número válido")

    # Preguntar cuántos formularios enviar
    max_submissions = max_submissions_arg or os.environ.get("MAX_SUBMISSIONS")
    if max_submissions:
        try:
            max_submissions = max(1, int(max_submissions))
        except Exception:
            max_submissions = None
    
    if max_submissions is None:
        if non_interactive:
            max_submissions = None  # Sin límite en modo no interactivo
        else:
            while True:
                try:
                    raw = input("\n¿Cuántos formularios quieres enviar? (Enter para enviar todos los disponibles): ").strip()
                    if not raw:
                        max_submissions = None  # Sin límite
                        break
                    max_submissions = int(raw)
                    if max_submissions >= 1:
                        break
                    else:
                        print("Por favor ingresa un número mayor a 0")
                except ValueError:
                    print("Por favor ingresa un número válido")

    return url, num_threads, max_submissions, emails_file

# Parsear booleanos de variables de entorno
def _parse_bool_env(val, default=None):
    if val is None:
        return default
    s = str(val).strip().lower()
    if s in ("1", "true", "yes", "y", "on"):
        return True
    if s in ("0", "false", "no", "n", "off"):
        return False
    return default

def load_emails():
    """Lee los correos disponibles y excluye aquellos que ya fueron utilizados."""
    try:
        with open(EMAILS_FILE, 'r', encoding='utf-8') as f:
            lst = [l.strip() for l in f if l.strip() and '@' in l.strip()]
    except FileNotFoundError:
        print(f"ERROR: No se encontró el archivo {EMAILS_FILE}")
        print("Asegúrate de proporcionar un archivo válido con correos (uno por línea).")
        sys.exit(1)
    
    # excluir usados
    used = set()
    if os.path.exists(used_path):
        with open(used_path, 'r', encoding='utf-8') as f:
            used = {l.strip() for l in f if l.strip()}
    remaining = [e for e in lst if e not in used]
    return remaining

# (La carga de correos y la creación de la cola ahora se realiza en initialize_emails())

# ===== Navegador helpers =====

def create_driver():
    """Inicializa un WebDriver de Chrome con las opciones necesarias para el bot."""
    opts = webdriver.ChromeOptions()
    # opts.add_argument("--headless=new")
    opts.add_argument("--start-maximized")
    # Evitar bloqueos por automatización en algunas builds
    opts.add_experimental_option("excludeSwitches", ["enable-automation"]) 
    opts.add_experimental_option('useAutomationExtension', False)
    # Reutilizar perfil de usuario si está configurado para evitar inicios de sesión
    user_data_dir = os.environ.get('CHROME_USER_DATA_DIR')
    profile_dir = os.environ.get('CHROME_PROFILE_DIR')
    tmp_dir = None
    if user_data_dir:
        # El usuario solicitó usar un perfil concreto (posible login ya hecho)
        opts.add_argument(f"--user-data-dir={user_data_dir}")
        if profile_dir:
            opts.add_argument(f"--profile-directory={profile_dir}")
    else:
        # Para evitar errores de Chrome con el perfil por defecto/en uso, usar un dir temporal único
        tmp_dir = tempfile.mkdtemp(prefix="chrome-selenium-")
        opts.add_argument(f"--user-data-dir={tmp_dir}")
    # Mejor estabilidad
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")
    opts.add_argument("--remote-allow-origins=*")
    if os.environ.get('CHROME_DETACH') == '1':
        opts.add_experimental_option("detach", True)
    drv = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    drv.set_page_load_timeout(40)
    w = WebDriverWait(drv, 20)
    # Limpiar el directorio temporal al salir
    if tmp_dir:
        def _cleanup_tmp():
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass
        atexit.register(_cleanup_tmp)
    return drv, w


def ensure_active_window():
    """Garantiza que exista y esté seleccionada una ventana manejable.
    Si la ventana actual fue cerrada (NoSuchWindowException), intenta
    cambiar a otra handle disponible. Lanza la excepción si no hay ninguna.
    """
    try:
        # tocar título para validar
        _ = driver.title  # noqa: F841
        return
    except NoSuchWindowException:
        pass
    except Exception:
        # otras excepciones no implican ventana cerrada
        return
    # Window parece cerrada: intentar cambiar a otra existente
    handles = []
    try:
        handles = driver.window_handles
    except Exception:
        handles = []
    if not handles:
        raise NoSuchWindowException("No hay ventanas de navegador activas")
    # Priorizar la última (suele ser principal)
    for h in reversed(handles):
        try:
            driver.switch_to.window(h)
            _ = driver.title
            return
        except Exception:
            continue
    # si ninguna funcionó
    raise NoSuchWindowException("No se pudo recuperar una ventana activa")


def open_form(url):
    """Abre el formulario objetivo manejando redirecciones y exigencias de inicio de sesión."""
    print("Navegando a:", url)
    try:
        driver.get(url)
    except TimeoutException:
        print("Timeout cargando, reintento #1...")
        driver.get(url)
    # Verificar si quedó en data:
    cur = ''
    try:
        cur = driver.current_url
        print("current_url:", cur)
    except Exception:
        pass
    if not cur or cur.startswith('data:'):
        # Intento 2: página neutra y luego al formulario
        try:
            driver.get('about:blank'); time.sleep(0.3)
            driver.get(url)
            cur = driver.current_url
            print("current_url tras about:blank:", cur)
        except Exception:
            pass
    if not cur or cur.startswith('data:'):
        # Intento 3: navegar a google.com y luego al form
        try:
            driver.get('https://www.google.com'); time.sleep(0.8)
            driver.get(url)
            cur = driver.current_url
            print("current_url tras google.com:", cur)
        except Exception:
            pass
    try:
        driver.save_screenshot(os.path.join(os.path.dirname(__file__), 'page_init.png'))
    except Exception:
        pass
    # Detectar redirección a login de Google y permitir login manual si está habilitado
    try:
        cur_lower = (driver.current_url or '').lower()
    except Exception:
        cur_lower = ''
    if any(p in cur_lower for p in ["accounts.google.com", "signin", "service=wise", "continue=https://docs.google.com"]):
        print("Parece que Google está pidiendo iniciar sesión para abrir el formulario.")
        if ALLOW_MANUAL_LOGIN:
            print("Inicia sesión manualmente en la ventana de Chrome y espera a que cargue el formulario.\n"
                  "Consejo: usa tu perfil de Chrome ya logueado con CHROME_USER_DATA_DIR/CHROME_PROFILE_DIR.")
            deadline = time.time() + 180
            last_notice = 0
            while time.time() < deadline:
                try:
                    cur = (driver.current_url or '').lower()
                except Exception:
                    cur = ''
                # ¿Ya estamos en docs.google.com/forms o se ve un <form>?
                try:
                    forms = driver.find_elements(By.TAG_NAME, 'form')
                    has_form = any(f.is_displayed() for f in forms)
                except Exception:
                    has_form = False
                if ('docs.google.com' in cur and 'forms' in cur) or has_form:
                    print("Login completado. Formulario visible.")
                    try:
                        ensure_active_window()
                    except Exception:
                        pass
                    break
                # De vez en cuando, intentar navegar a la URL objetivo otra vez
                if time.time() - last_notice > 15:
                    last_notice = time.time()
                    try:
                        driver.get(url)
                    except Exception:
                        pass
                time.sleep(1)
            else:
                print("No se completó el inicio de sesión a tiempo (180s).")
                return False
        else:
            print("Sugerencias: usa un formulario público o ejecuta con un perfil de Chrome ya iniciado (ver variables CHROME_USER_DATA_DIR/CHROME_PROFILE_DIR).")
            return False

    try:
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "form")))
        # Revisar si aparece el modal interno de 'Iniciar sesión para continuar'
        try:
            modals = driver.find_elements(By.XPATH, "//div[contains(., 'Iniciar sesión para continuar') or contains(., 'Sign in to continue')]//ancestor::div[@role='dialog']")
            visible_modals = [m for m in modals if m.is_displayed()]
            if visible_modals:
                print("El formulario muestra un modal que exige iniciar sesión.")
                if ALLOW_MANUAL_LOGIN:
                    print("Voy a pulsar 'Iniciar sesión' si es visible. Luego inicia sesión manualmente y volveré al formulario.")
                    # Intentar pulsar el botón "Iniciar sesión" dentro del modal
                    try:
                        login_btn = None
                        # Buscar dentro del modal visible
                        for m in visible_modals:
                            try:
                                cand = m.find_elements(By.XPATH, ".//*[@role='button' or self::button][contains(., 'Iniciar sesión') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sign in')]")
                                cand = [c for c in cand if c.is_displayed()]
                                if cand:
                                    login_btn = cand[0]
                                    break
                            except Exception:
                                continue
                        if login_btn is None:
                            # búsqueda global como respaldo
                            cand = driver.find_elements(By.XPATH, "//*[@role='button' or self::button][contains(., 'Iniciar sesión') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sign in')]")
                            cand = [c for c in cand if c.is_displayed()]
                            if cand:
                                login_btn = cand[0]
                        if login_btn is not None:
                            try:
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", login_btn)
                            except Exception:
                                pass
                            try:
                                login_btn.click()
                            except Exception:
                                try:
                                    ActionChains(driver).move_to_element(login_btn).pause(0.05).click().perform()
                                except Exception:
                                    driver.execute_script("arguments[0].click();", login_btn)
                            print("Botón 'Iniciar sesión' pulsado. Completa el login en la ventana.")
                    except Exception:
                        pass

                    print("Por favor, inicia sesión manualmente en la ventana de Chrome y vuelve al formulario. Detectaré cuando desaparezca el modal.")
                    # Esperar hasta 300s a que el modal desaparezca y aparezca el form
                    deadline = time.time() + 300
                    while time.time() < deadline:
                        try:
                            forms = driver.find_elements(By.TAG_NAME, 'form')
                            modals = driver.find_elements(By.XPATH, "//div[@role='dialog']")
                            if forms and not any(m.is_displayed() for m in modals):
                                print("Login completado y formulario visible.")
                                try:
                                    ensure_active_window()
                                except Exception:
                                    pass
                                break
                        except Exception:
                            pass
                        time.sleep(1)
                    else:
                        print("No se completó el inicio de sesión a tiempo.")
                        return False
                else:
                    print("Usa tu perfil de Chrome ya logueado (CHROME_USER_DATA_DIR/CHROME_PROFILE_DIR) o cambia la configuración del formulario a público.")
                    return False
        except Exception:
            pass
        return True
    except TimeoutException:
        print("No se detectó el <form> tras abrir la URL.")
        try:
            print("current_url final:", driver.current_url)
            driver.save_screenshot(os.path.join(os.path.dirname(__file__), 'page_load_timeout.png'))
        except Exception:
            pass
        return False

# ===== Helpers =====

def fill_email_field(driver):
    email_val = CURRENT_EMAIL
    # 1) Por aria-label (más fiable)
    candidates = driver.find_elements(By.CSS_SELECTOR, "input[aria-label], textarea[aria-label]")
    for el in candidates:
        try:
            label = (el.get_attribute("aria-label") or "").strip().lower()
            if ("correo" in label) or ("email" in label) or ("e-mail" in label):
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                el.clear()
                el.send_keys(email_val)
                print(f"Correo llenado: {email_val}")
                return True
        except Exception:
            continue
    # 2) Por tipo de input email
    try:
        el = driver.find_element(By.CSS_SELECTOR, "input[type='email']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
        el.clear(); el.send_keys(email_val)
        print(f"Correo llenado (type=email): {email_val}")
        return True
    except Exception:
        pass
    print("No se encontró campo de correo")
    return False


def fill_all_radio_groups():
    """Selecciona una opción aleatoria en cada grupo de radio visible en la sección actual."""
    try:
        groups = driver.find_elements(By.CSS_SELECTOR, "div[role='radiogroup']")
        # Aleatorizar orden de grupos
        groups = [g for g in groups if g.is_displayed()]
        random.shuffle(groups)
        any_clicked = False
        for grp in groups:
            radios = [r for r in grp.find_elements(By.CSS_SELECTOR, "div[role='radio']") if r.is_displayed()]
            if not radios:
                continue
            # Si ya hay seleccionado, saltar
            already = [r for r in radios if r.get_attribute("aria-checked") == "true"]
            if already:
                continue
            random.shuffle(radios)
            choice = radios[0]
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", choice)
                human_pause()
                choice.click()
                any_clicked = True
            except Exception:
                try:
                    driver.execute_script("arguments[0].click();", choice)
                    any_clicked = True
                except Exception:
                    pass
        return any_clicked
    except Exception as e:
        print(f"Error al seleccionar radios: {e}")
        return False


def select_random_options_in_all_dropdowns():
    """Selecciona una opción aleatoria válida en cada dropdown visible de la sección actual.
    Estrategia:
    1) Detectar triggers típicos del dropdown
    2) Abrir con click/JS/teclado y esperar el listbox visible
    3) Elegir una opción aleatoria de ese listbox y hacer click
    """
    trigger_selectors = [
        "div[aria-haspopup='listbox']",
        "div[role='combobox']",
        "div[role='listbox'][tabindex]"
    ]

    def find_triggers():
        results = []
        for sel in trigger_selectors:
            try:
                ensure_active_window()
            except Exception:
                return []
            for el in driver.find_elements(By.CSS_SELECTOR, sel):
                try:
                    if el.is_displayed():
                        results.append(el)
                except Exception:
                    continue
        # Aleatorizar orden de triggers
        random.shuffle(results)
        return results

    def open_dropdown(trigger):
        opened = False
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", trigger)
        except Exception:
            pass
        try:
            ActionChains(driver).move_to_element(trigger).pause(0.05).click().perform()
            opened = True
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", trigger)
                opened = True
            except Exception:
                pass
        if not opened:
            try:
                driver.execute_script("arguments[0].focus();", trigger)
                trigger.send_keys(Keys.ENTER)
                opened = True
            except Exception:
                pass
        return opened

    def get_visible_listbox():
        try:
            ensure_active_window()
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='listbox']"))
            )
        except Exception:
            return None
        listboxes = driver.find_elements(By.CSS_SELECTOR, "div[role='listbox']")
        visibles = [lb for lb in listboxes if lb.is_displayed()]
        return visibles[-1] if visibles else None

    def choose_random_option_from_listbox(listbox):
        try:
            options = listbox.find_elements(By.CSS_SELECTOR, "div[role='option']")
        except Exception:
            return False
        # Aleatorizar orden
        random.shuffle(options)
        for opt in options:
            try:
                if not opt.is_displayed():
                    continue
                txt = (opt.text or "").strip()
                if not txt:
                    continue
                if txt.lower() in {"elige", "choose", "select"}:
                    continue
                if opt.get_attribute("aria-disabled") == "true":
                    continue
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", opt)
                human_pause()
                try:
                    ActionChains(driver).move_to_element(opt).pause(0.05).click().perform()
                except Exception:
                    driver.execute_script("arguments[0].click();", opt)
                return True
            except Exception:
                continue
        return False

    triggers = find_triggers()
    print(f"Dropdowns visibles (detectados): {len(triggers)}")
    any_selected = False

    for trig in triggers:
        opened = open_dropdown(trig)
        selected = False
        if opened:
            lb = get_visible_listbox()
            if lb:
                selected = choose_random_option_from_listbox(lb)
        if not selected:
            try:
                # Fallback completamente aleatorio por teclado
                focus_el = trig
                try:
                    focus_el = trig.find_element(By.CSS_SELECTOR, "[role='combobox'], input, [tabindex]")
                except Exception:
                    pass
                driver.execute_script("arguments[0].focus();", focus_el)
                focus_el.send_keys(Keys.ENTER)
                for _ in range(random.randint(1, 10)):
                    focus_el.send_keys(Keys.ARROW_DOWN)
                    human_pause(0.03, 0.09)
                focus_el.send_keys(Keys.ENTER)
                selected = True
            except Exception:
                selected = False
        any_selected = any_selected or selected
        human_pause(0.08, 0.2)

    return any_selected


def fill_all_checkboxes():
    """Marca al menos una opción en cada grupo de checkboxes visible.
    Heurística simple: recorre los checkboxes visibles y, por cada contenedor
    más cercano con rol de grupo/list, marca el primero que no esté marcado.
    """
    try:
        checkboxes = [c for c in driver.find_elements(By.CSS_SELECTOR, "div[role='checkbox']") if c.is_displayed()]
    except Exception:
        return False

    any_clicked = False
    seen_containers = set()
    for c in checkboxes:
        try:
            container = None
            try:
                container = c.find_element(By.XPATH, "ancestor::*[@role='group' or @role='list' or @role='radiogroup'][1]")
            except Exception:
                try:
                    container = c.find_element(By.XPATH, "ancestor::*[contains(@class,'freebirdFormviewerComponentsQuestionBaseRoot')][1]")
                except Exception:
                    container = None
            key = id(container) if container else id(c)
            if key in seen_containers:
                continue
            checked = (c.get_attribute('aria-checked') == 'true')
            if not checked:
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", c)
                except Exception:
                    pass
                try:
                    c.click()
                except Exception:
                    try:
                        ActionChains(driver).move_to_element(c).pause(0.05).click().perform()
                    except Exception:
                        try:
                            driver.execute_script("arguments[0].click();", c)
                        except Exception:
                            continue
                any_clicked = True
            seen_containers.add(key)
        except Exception:
            continue
    return any_clicked


def _random_name():
    first = random.choice(["Juan", "Luis", "Ana", "María", "Pedro", "Laura", "Carlos", "Sofía"]) 
    last = random.choice(["Pérez", "Gómez", "Rodríguez", "López", "García", "Hernández"]) 
    return f"{first} {last}"


def _random_phone():
    return ''.join(random.choice('6789') if i == 0 else random.choice('0123456789') for i in range(10))


def _random_date():
    try:
        start = datetime.date.today() - datetime.timedelta(days=3650)
        end = datetime.date.today()
        delta = (end - start).days
        d = start + datetime.timedelta(days=random.randint(0, max(1, delta)))
        return d.strftime('%d/%m/%Y')
    except Exception:
        return '01/01/2000'


def _random_time():
    try:
        h = random.randint(8, 20)
        m = random.choice([0, 15, 30, 45])
        return f"{h:02d}:{m:02d}"
    except Exception:
        return '10:00'


def fill_all_text_inputs():
    """Rellena inputs/textarea visibles con valores plausibles para permitir enviar.
    Usa heurísticas por aria-label/texto para elegir el formato.
    """
    try:
        fields = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
    except Exception:
        return False
    any_filled = False
    for el in fields:
        try:
            if not el.is_displayed():
                continue
            if el.get_attribute('disabled') or el.get_attribute('readonly'):
                continue
            val = (el.get_attribute('value') or '').strip()
            if val:
                continue
            label = (el.get_attribute('aria-label') or '').strip().lower()
            # Buscar hint por elementos cercanos
            if not label:
                try:
                    near = el.find_element(By.XPATH, "ancestor::*[1]//label|ancestor::*[2]//label|ancestor::*[2]//*[self::span or self::div][normalize-space()][1]")
                    label = (near.text or '').strip().lower()
                except Exception:
                    label = ''

            to_fill = None
            if any(k in label for k in ['correo', 'email', 'e-mail', 'mail']):
                to_fill = CURRENT_EMAIL
            elif any(k in label for k in ['nombre', 'name']):
                to_fill = _random_name()
            elif any(k in label for k in ['apellido', 'last name', 'apellidos']):
                to_fill = _random_name().split(' ')[1]
            elif any(k in label for k in ['dni', 'cédula', 'cedula', 'documento', 'id']):
                to_fill = ''.join(random.choice('0123456789') for _ in range(8))
            elif any(k in label for k in ['tel', 'cel', 'whatsapp', 'phone']):
                to_fill = _random_phone()
            elif any(k in label for k in ['edad', 'age']):
                to_fill = str(random.randint(18, 60))
            elif any(k in label for k in ['fecha', 'date']):
                to_fill = _random_date()
            elif any(k in label for k in ['hora', 'time']):
                to_fill = _random_time()
            else:
                to_fill = random.choice(['Sí', 'Ok', 'N/A', 'Acepto', _random_name()])

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
            human_pause(0.05, 0.12)
            try:
                el.clear()
            except Exception:
                pass
            el.send_keys(to_fill)
            any_filled = True
        except Exception:
            continue
    return any_filled


def dump_debug_info(context=""):
    try:
        title = driver.title
    except Exception:
        title = "(sin título)"
    print(f"\n[DEBUG] Estado de página {context} | title: {title}")
    try:
        btns = driver.find_elements(By.XPATH, "//*[@role='button' or self::button or self::a]")
        visibles = [b for b in btns if b.is_displayed()]
        print(f"[DEBUG] Botones visibles: {len(visibles)}")
        for b in visibles[:10]:
            try:
                txt = (b.text or b.get_attribute('aria-label') or '').strip().replace('\n',' ')
                print("[DEBUG]   -", txt[:120])
            except Exception:
                continue
    except Exception:
        pass
    try:
        radios = driver.find_elements(By.CSS_SELECTOR, "div[role='radiogroup']")
        print(f"[DEBUG] Radiogrupos: {len(radios)}")
    except Exception:
        pass
    try:
        lists = driver.find_elements(By.CSS_SELECTOR, "div[role='listbox']")
        print(f"[DEBUG] Listbox presentes: {len(lists)}")
    except Exception:
        pass

# ...existing code...

def click_button_by_text(texto):
    """Busca y hace clic en un botón por texto de forma robusta.
    Especializado para Google Forms: prioriza el botón del pie con "Enviar".
    """
    texto = (texto or '').strip().lower()
    synonyms = {
        'next': ['siguiente', 'continuar', 'next', 'continue'],
        'submit': ['enviar', 'enviar respuestas', 'enviar respuesta', 'submit', 'send']
    }
    candidates = [texto]
    if texto in ('siguiente', 'next', 'continuar', 'continue'):
        candidates = synonyms['next']
    elif texto in ('enviar', 'submit', 'send', 'enviar respuesta', 'enviar respuestas'):
        candidates = synonyms['submit']

    # Forzar render: scroll al inicio y al final
    try:
        driver.execute_script("window.scrollTo(0, 0);")
        human_pause(0.05, 0.12)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        human_pause(0.05, 0.12)
    except Exception:
        pass

    # Construir lista de XPaths a evaluar en bloque (sin esperar a 'clickable')
    xpaths = []
    for t in candidates:
        xpaths.extend([
            # role=button que contenga el texto en spans o en el propio nodo
            f"//div[@role='button'][.//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')] or contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')]",
            # botón HTML y enlaces
            f"//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')]",
            f"//a[@role='button' and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')]",
            # por aria-label
            f"//*[@role='button' and contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')]",
        ])

    def _score_button(el):
        try:
            rect = el.rect or {}
            y = rect.get('y', 0)
        except Exception:
            y = 0
        try:
            label = (el.text or el.get_attribute('aria-label') or '').strip().lower()
        except Exception:
            label = ''
        score = 0
        if 'enviar' in label:
            score += 50
        if 'submit' in label or 'send' in label:
            score += 30
        # Más cerca del final de la página, mayor puntuación
        score += min(2000, y)
        # Evitar botones de "Atrás" o limpiar
        if any(x in label for x in ['atrás', 'anterior', 'borrar', 'limpiar', 'clear', 'back']):
            score -= 1000
        # Evitar deshabilitados
        try:
            if el.get_attribute('aria-disabled') == 'true' or el.get_attribute('disabled'):
                score -= 1000
        except Exception:
            pass
        return score, y

    # Recolectar candidatos visibles
    found = []
    for xp in xpaths:
        try:
            elems = driver.find_elements(By.XPATH, xp)
        except Exception:
            elems = []
        for el in elems:
            try:
                if el.is_displayed():
                    s, y = _score_button(el)
                    found.append((s, y, el))
            except Exception:
                continue

    # Si no se encontraron por texto, intentar por estructura del pie del formulario
    if not found:
        try:
            form = driver.find_element(By.CSS_SELECTOR, 'form')
            # Botones Material de Google Forms en el pie
            guess = [b for b in form.find_elements(By.CSS_SELECTOR, "div[role='button'], button, a[role='button']") if b.is_displayed()]
            for b in guess:
                s, y = _score_button(b)
                found.append((s, y, b))
        except Exception:
            pass

    if not found:
        dump_debug_info(context=f"(no se encontró '{texto}')")
        return False

    # Ordenar por score descendente y posición Y (más abajo primero)
    found.sort(key=lambda t: (t[0], t[1]), reverse=True)

    # Probar múltiples estrategias de clic para el mejor candidato
    for _, _, btn in found[:5]:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
        except Exception:
            pass
        human_pause(0.05, 0.12)
        # 1) click directo
        try:
            btn.click()
            return True
        except Exception:
            pass
        # 2) ActionChains
        try:
            ActionChains(driver).move_to_element(btn).pause(0.05).click().perform()
            return True
        except Exception:
            pass
        # 3) JS click
        try:
            driver.execute_script("arguments[0].click();", btn)
            return True
        except Exception:
            pass
        # 4) Click por coordenadas usando elementFromPoint
        try:
            rect = driver.execute_script(
                "const r=arguments[0].getBoundingClientRect(); return {x: r.left + r.width/2, y: r.top + r.height/2};",
                btn,
            )
            if rect and 'x' in rect and 'y' in rect:
                driver.execute_script(
                    "var el=document.elementFromPoint(arguments[0], arguments[1]); if(el){el.click();}",
                    int(rect['x']), int(rect['y'])
                )
                return True
        except Exception:
            pass

    dump_debug_info(context=f"(no se pudo hacer clic en '{texto}')")
    return False


def wait_for_submission_confirmation(timeout=15):
    # 1) Confirmación por URL de Google Forms: termina en /formResponse
    try:
        WebDriverWait(driver, timeout).until(lambda d: 'formResponse' in (d.current_url or ''))
        return True
    except Exception:
        pass

    # 2) Confirmación por textos genéricos (muy permisivo por mensajes personalizados)
    texts = [
        "se ha registrado tu respuesta",
        "tu respuesta ha sido registrada",
        "respuesta registrada",
        "your response has been recorded",
        "response recorded",
        "gracias",
        "thanks",
        "thank you",
        "se envió tu respuesta",
        "respuesta enviada"
    ]
    try:
        wait2 = WebDriverWait(driver, timeout)
        xp_parts = [
            f"contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')"
            for t in texts
        ]
        xp = " or ".join(xp_parts)
        confirmation_xpath = (
            f"//div[@role='heading'][{xp}] | "
            f"//body//*[self::h1 or self::h2 or self::h3 or self::div or self::span or self::p][{xp}]"
        )
        wait2.until(EC.presence_of_element_located((By.XPATH, confirmation_xpath)))
        return True
    except Exception:
        return False

# Reset a formulario vacío desde la pantalla de confirmación

def reset_form_or_navigate():
    if click_button_by_text("Enviar otra respuesta") or click_button_by_text("Submit another response"):
        return True
    try:
        driver.get(FORM_URL)
        return True
    except Exception:
        return False

def click_submit_by_image_dom():
    """Intenta localizar un botón de envío representado por una imagen.
    Criterios:
      - input[type=image]
      - img con alt/src que incluya 'enviar' o 'submit'
      - sube a ancestro clickable si existe
    """
    selectors = [
        "input[type='image']",
        "img[alt]",
        "img[src]",
    ]
    words = ["enviar", "submit", "send"]
    try:
        # Recolectar candidatos visibles con score
        candidates = []
        for sel in selectors:
            for el in driver.find_elements(By.CSS_SELECTOR, sel):
                try:
                    if not el.is_displayed():
                        continue
                    tag = el.tag_name.lower()
                    alt = (el.get_attribute('alt') or '').strip().lower()
                    src = (el.get_attribute('src') or '').strip().lower()
                    ok = False
                    if tag == 'input' and (el.get_attribute('type') or '').lower() == 'image':
                        ok = True
                    if any(w in alt or w in src for w in words):
                        ok = True
                    if not ok:
                        continue
                    # Scoring: priorizar enviar.png / src que contenga 'enviar'
                    score = 0
                    if 'enviar.png' in src:
                        score += 50
                    if 'enviar' in src or 'enviar' in alt:
                        score += 20
                    if tag == 'input':
                        score += 10
                    # Más abajo en la página suele ser el botón final
                    try:
                        ypos = el.rect.get('y', 0)
                    except Exception:
                        ypos = 0
                    candidates.append((score, ypos, el))
                except Exception:
                    continue
        # Orden: score desc, y desc
        candidates.sort(key=lambda t: (t[0], t[1]), reverse=True)

        for _, _, el in candidates:
            try:
                # Subir al ancestro clickable si corresponde
                target = el
                try:
                    target = el.find_element(By.XPATH, "ancestor::*[@role='button' or self::button or self::a][1]")
                except Exception:
                    target = el
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
                human_pause(0.05, 0.12)
                try:
                    target.click()
                except Exception:
                    try:
                        ActionChains(driver).move_to_element(target).pause(0.05).click().perform()
                    except Exception:
                        driver.execute_script("arguments[0].click();", target)
                return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def attempt_submit():
    """Intenta enviar el formulario por múltiples estrategias."""
    # 0) Búsqueda directa de enviar.png (prioridad solicitada)
    if click_enviar_png():
        return True
    # 1) Por botones de texto conocidos
    if click_button_by_text("enviar") or click_button_by_text("submit") or click_button_by_text("send") or click_button_by_text("enviar respuesta") or click_button_by_text("enviar respuestas"):
        return True
    # 2) Por imagen (input[type=image] o img alt/src)
    if click_submit_by_image_dom():
        return True
    # 2.5) Fuerza por JavaScript buscando cualquier nodo con texto "Enviar" y clic en ancestro clickable
    if force_click_submit_js():
        return True
    # 3) Enter en el formulario como último recurso
    try:
        form = driver.find_element(By.CSS_SELECTOR, 'form')
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", form)
        form.send_keys(Keys.ENTER)
        return True
    except Exception:
        pass
    return False


def click_enviar_png():
    """Encuentra y pulsa específicamente una imagen cuyo src contenga 'enviar.png'."""
    try:
        candidates = []
        # 1) IMG o INPUT[IMAGE] cuyo SRC contenga 'enviar.png' o 'enviar'
        xp = (
            "//img[contains(translate(@src, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'enviar.png') or "
            "contains(translate(@src, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'enviar')] | "
            "//input[@type='image' and (contains(translate(@src, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'enviar.png') or "
            "contains(translate(@src, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'enviar'))]"
        )
        for el in driver.find_elements(By.XPATH, xp):
            try:
                if el.is_displayed():
                    candidates.append((80 if 'enviar.png' in (el.get_attribute('src') or '').lower() else 60, el))
            except Exception:
                continue

        # 2) Cualquier elemento cuyo style inline contenga 'enviar' (posible background-image)
        style_elems = driver.find_elements(By.XPATH, "//*[contains(translate(@style, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'enviar')]")
        for el in style_elems:
            try:
                if el.is_displayed():
                    candidates.append((40, el))
            except Exception:
                continue

        # 3) Computed style background-image conteniendo 'enviar'
        try:
            js = (
                "return Array.from(document.querySelectorAll('*')).filter(e=>{"
                "try{"
                "const bg=(getComputedStyle(e).backgroundImage||'');"
                "return String(bg).toLowerCase().includes('enviar');"
                "}catch(err){return false;}"
                "}).slice(-200);"
            )  # limitar para no devolver demasiados
            bg_elems = driver.execute_script(js) or []
            for el in bg_elems:
                try:
                    if el.is_displayed():
                        candidates.append((50, el))
                except Exception:
                    continue
        except Exception:
            pass

        if not candidates:
            # Fallback: búsqueda visual en la página con captura + OpenCV
            if click_enviar_png_by_screenshot():
                return True
            return False

        # Ordenar por score y posición Y descendente
        scored = []
        for score, el in candidates:
            try:
                y = el.rect.get('y', 0)
                scored.append((score, y, el))
            except Exception:
                continue
        scored.sort(key=lambda t: (t[0], t[1]), reverse=True)

        for _, _, el in scored:
            try:
                target = el
                try:
                    target = el.find_element(By.XPATH, "ancestor::*[@role='button' or self::button or self::a][1]")
                except Exception:
                    target = el
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
                human_pause(0.05, 0.12)
                try:
                    target.click()
                except Exception:
                    try:
                        ActionChains(driver).move_to_element(target).pause(0.05).click().perform()
                    except Exception:
                        driver.execute_script("arguments[0].click();", target)
                return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def click_enviar_png_by_screenshot(threshold=0.82):
    """Hace scroll por la página, toma capturas y usa OpenCV para localizar
    la imagen enviar.png en pantalla; luego hace clic en el centro encontrado.
    Requiere opencv-python y numpy. Si no están, retorna False sin error.
    """
    img_path = os.path.join(os.path.dirname(__file__), 'enviar.png')
    if not os.path.exists(img_path):
        print("No se encontró 'enviar.png' en la carpeta del script.")
        return False
    try:
        import cv2  # type: ignore
        import numpy as np  # type: ignore
    except Exception:
        print("OpenCV/Numpy no instalados; omito búsqueda visual de 'enviar.png'.")
        return False

    template = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)
    if template is None:
        print("No se pudo leer 'enviar.png' como imagen.")
        return False
    th, tw = template.shape[:2]

    try:
        dpr = driver.execute_script("return window.devicePixelRatio") or 1
    except Exception:
        dpr = 1

    try:
        scroll_h = int(driver.execute_script("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"))
        inner_h = int(driver.execute_script("return window.innerHeight"))
    except Exception:
        scroll_h, inner_h = 0, 0

    max_y = max(0, scroll_h - inner_h)
    step = max(1, int(inner_h * 0.85)) if inner_h else 800

    # Intentar en múltiples posiciones de scroll
    positions = list(range(0, max_y + 1, step)) or [0]
    for y in positions:
        try:
            driver.execute_script("window.scrollTo(0, arguments[0]);", y)
        except Exception:
            pass
        # Sin pausas - continuar inmediatamente
        time.sleep(0.05)  # Mínima pausa técnica

        try:
            png = driver.get_screenshot_as_png()
        except Exception:
            continue
        image = cv2.imdecode(np.frombuffer(png, np.uint8), cv2.IMREAD_GRAYSCALE)
        if image is None:
            continue

        found = False
        loc = (0, 0)
        tsize = (tw, th)
        for scale in [0.85, 0.9, 0.95, 1.0, 1.05, 1.1, 1.15]:
            try:
                scaled = cv2.resize(template, (max(1, int(tw * scale)), max(1, int(th * scale))),
                                    interpolation=cv2.INTER_AREA if scale < 1.0 else cv2.INTER_CUBIC)
                res = cv2.matchTemplate(image, scaled, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(res)
                if max_val >= threshold:
                    found = True
                    loc = max_loc
                    tsize = (scaled.shape[1], scaled.shape[0])
                    break
            except Exception:
                continue

        if not found:
            continue

        cx = loc[0] + tsize[0] / 2.0
        cy = loc[1] + tsize[1] / 2.0
        css_x = int(cx / dpr)
        css_y = int(cy / dpr)

        try:
            body = driver.find_element(By.TAG_NAME, 'body')
        except Exception:
            body = None

        try:
            if body is not None:
                ActionChains(driver).move_to_element_with_offset(body, css_x, css_y).click().perform()
            else:
                driver.execute_script(
                    "var el=document.elementFromPoint(arguments[0], arguments[1]); if(el){el.click();}", css_x, css_y
                )
            return True
        except Exception:
            try:
                driver.execute_script(
                    "var el=document.elementFromPoint(arguments[0], arguments[1]); if(el){el.click();}", css_x, css_y
                )
                return True
            except Exception:
                continue

    return False


def force_click_submit_js():
        """Ultimo recurso: usa JavaScript para localizar por texto 'Enviar'/'Submit'/'Send'
        en cualquier nodo visible, y hace clic sobre su ancestro 'button-like'.
        Devuelve True si pudo ejecutar algún clic.
        """
        try:
                js = r"""
                (function(){
                    function visible(e){
                        try{
                            const st=getComputedStyle(e); const r=e.getBoundingClientRect();
                            return st.display!=='none' && st.visibility!=='hidden' && r.width>4 && r.height>4;
                        }catch(_){return false}
                    }
                    function clickableAncestor(e){
                        const sel = "button, [role=button], a[role=button], .appsMaterialWizButtonEl, .quantumWizButtonPaperbuttonEl, .freebirdFormviewerViewNavigationSubmitButton";
                        let n=e; let depth=0;
                        while(n && depth<6){
                            if(n.matches && n.matches(sel)) return n;
                            n=n.parentElement; depth++;
                        }
                        return e;
                    }
                    const texts=["enviar","enviar respuestas","enviar respuesta","submit","send"]; 
                    const all=Array.from(document.querySelectorAll('*'));
                    const cand=[];
                    for(const el of all){
                        try{
                            if(!visible(el)) continue;
                            const t=(el.textContent||'').toLowerCase().trim();
                            if(!t) continue;
                            if(texts.some(x=>t.includes(x))){
                                const r=el.getBoundingClientRect();
                                cand.push({el, y:r.top});
                            }
                        }catch(_){continue}
                    }
                    if(!cand.length) return false;
                    cand.sort((a,b)=>a.y-b.y); // de arriba a abajo
                    // probar desde los de más abajo (pie de página)
                    for(let i=cand.length-1;i>=0;i--){
                        const target=clickableAncestor(cand[i].el);
                        try{
                            target.scrollIntoView({block:'center'});
                            const evt = new MouseEvent('click', {bubbles:true,cancelable:true,view:window});
                            if(target.dispatchEvent(evt)){
                                if(typeof target.click==='function') target.click();
                                return true;
                            }
                        }catch(_){continue}
                    }
                    return false;
                })();
                """
                ok = driver.execute_script(js)
                return bool(ok)
        except Exception:
                return False

# Proceso de una sola respuesta
def process_one_submission():
    """Procesa un solo envío con el driver específico"""
    global CURRENT_EMAIL
    email = CURRENT_EMAIL
    try:
        ensure_active_window()
    except Exception as e:
        print(f"[{email}] Ventana inactiva al iniciar: {e}")
        return False

    fill_email_field(driver)
    fill_all_text_inputs()
    select_random_options_in_all_dropdowns()
    fill_all_radio_groups()
    fill_all_checkboxes()

    for _ in range(9):
        try:
            ensure_active_window()
        except Exception as e:
            print(f"[{email}] Ventana inactiva antes de avanzar: {e}")
            return False
        if click_button_by_text("siguiente"):
            fill_email_field(driver)
            fill_all_text_inputs()
            select_random_options_in_all_dropdowns()
            fill_all_radio_groups()
            fill_all_checkboxes()
            continue
        elif click_button_by_text("continuar") or click_button_by_text("next") or click_button_by_text("continue"):
            fill_email_field(driver)
            fill_all_text_inputs()
            select_random_options_in_all_dropdowns()
            fill_all_radio_groups()
            fill_all_checkboxes()
            continue
        else:
            # Intentar enviar
            print(f"[{email}] Preparando envío...")
            try:
                ensure_active_window()
            except Exception as e:
                print(f"[{email}] Ventana inactiva al enviar: {e}")
                return False
            # Prioridad: imagen enviar.png, si no, otras estrategias
            if click_enviar_png() or attempt_submit():
                # Sin espera - máxima velocidad
                print(f"[{email}] Se pulsó Enviar.")
                return True
            else:
                print(f"[{email}] No se encontró botón para avanzar ni enviar.")
                dump_debug_info(context="(sin botón avanzar/enviar)")
                return False

    print(f"[{email}] Límite de secciones alcanzado sin enviar.")
    dump_debug_info(context="(límite secciones)")
    return False

# === Función para manejar un solo hilo ===
def worker_thread(thread_id):
    """Ejecución principal de un hilo de trabajo que procesa envíos del formulario."""
    global success_count
    thread_driver = None
    thread_wait = None
    
    try:
        # Crear navegador independiente para este hilo
        thread_driver, thread_wait = create_driver()
        print(f"[Hilo {thread_id}] Navegador iniciado")
        
        while True:
            try:
                # Obtener siguiente email de la cola thread-safe
                email = email_queue.get_nowait()
            except queue.Empty:
                print(f"[Hilo {thread_id}] No hay más correos. Terminando hilo.")
                break
            
            print(f"[Hilo {thread_id}] Procesando: {email}")
            
            # Abrir formulario
            if not open_form_threaded(FORM_URL, thread_driver, thread_wait):
                print(f"[Hilo {thread_id}] No se pudo abrir el formulario para {email}")
                email_queue.task_done()
                continue
            
            # Procesar formulario
            success = process_one_submission_threaded(email, thread_driver, thread_wait)
            
            if success:
                # Incrementar contador de éxito de forma thread-safe
                with success_lock:
                    success_count += 1
                    current_success = success_count
                
                # Marcar email como usado
                with open(used_path, 'a', encoding='utf-8') as f:
                    f.write(email + "\n")
                
                print(f"[Hilo {thread_id}] Envío completado con {email} ({current_success} totales)")
                
                # Verificar límite de envíos
                if MAX_SUBMISSIONS is not None and current_success >= MAX_SUBMISSIONS:
                    print(f"[Hilo {thread_id}] Límite alcanzado: {current_success}. Terminando hilo.")
                    email_queue.task_done()
                    break
            else:
                print(f"[Hilo {thread_id}] No fue posible enviar la respuesta con {email}")
            
            email_queue.task_done()
            
    except Exception as e:
        print(f"[Hilo {thread_id}] Error crítico: {e}")
    finally:
        # Cerrar navegador del hilo
        if thread_driver:
            try:
                thread_driver.quit()
                print(f"[Hilo {thread_id}] Navegador cerrado")
            except Exception:
                pass

def open_form_threaded(url, driver, wait):
    """Abre el formulario objetivo usando un WebDriver dedicado por hilo."""
    print(f"[Thread] Navegando a: {url}")
    try:
        driver.get(url)
    except TimeoutException:
        print("[Thread] Timeout cargando, reintento #1...")
        driver.get(url)
    
    # Verificar si quedó en data:
    cur = ''
    try:
        cur = driver.current_url
        print("[Thread] current_url:", cur)
    except Exception:
        pass
    
    if not cur or cur.startswith('data:'):
        # Intento 2: página neutra y luego al formulario
        try:
            driver.get('about:blank')
            time.sleep(0.3)
            driver.get(url)
            cur = driver.current_url
            print("[Thread] current_url tras about:blank:", cur)
        except Exception:
            pass
    
    if not cur or cur.startswith('data:'):
        # Intento 3: navegar a google.com y luego al form
        try:
            driver.get('https://www.google.com')
            time.sleep(0.8)
            driver.get(url)
            cur = driver.current_url
            print("[Thread] current_url tras google.com:", cur)
        except Exception:
            pass
    
    # Detectar redirección a login de Google
    try:
        cur_lower = (driver.current_url or '').lower()
    except Exception:
        cur_lower = ''
    
    if any(p in cur_lower for p in ["accounts.google.com", "signin", "service=wise", "continue=https://docs.google.com"]):
        print("[Thread] Parece que Google está pidiendo iniciar sesión para abrir el formulario.")
        if ALLOW_MANUAL_LOGIN:
            print("[Thread] Esperando login manual en esta ventana...")
            deadline = time.time() + 180
            while time.time() < deadline:
                try:
                    cur = (driver.current_url or '').lower()
                except Exception:
                    cur = ''
                # ¿Ya estamos en docs.google.com/forms o se ve un <form>?
                try:
                    forms = driver.find_elements(By.TAG_NAME, 'form')
                    has_form = any(f.is_displayed() for f in forms)
                except Exception:
                    has_form = False
                if ('docs.google.com' in cur and 'forms' in cur) or has_form:
                    print("[Thread] Login completado. Formulario visible.")
                    break
                # Reintento navegar cada 15s
                if time.time() % 15 < 1:
                    try:
                        driver.get(url)
                    except Exception:
                        pass
                time.sleep(1)
            else:
                print("[Thread] No se completó el inicio de sesión a tiempo (180s).")
                return False
        else:
            print("[Thread] Login requerido pero no permitido.")
            return False

    try:
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "form")))
        return True
    except Exception:
        print("[Thread] No se encontró formulario después del login")
        return False

def process_one_submission_threaded(email, driver, wait):
    """Gestiona el llenado y envío del formulario para un correo dentro de un hilo."""
    global CURRENT_EMAIL
    CURRENT_EMAIL = email
    
    try:
        # Asegurar ventana activa específica para este driver
        _ = driver.title
    except NoSuchWindowException:
        print(f"[{email}] Ventana cerrada inesperadamente")
        return False
    except Exception as e:
        print(f"[{email}] Error verificando ventana: {e}")
        return False

    for step in range(1, 21):  # Máximo 20 secciones
        print(f"[{email}] Sección {step}")
        
        # Llenar campos en orden iniciando por el correo
        email_filled = fill_email_field_threaded(driver, email, wait)
        if not email_filled:
            email_filled = _force_fill_email_js(driver, email)
        if not email_filled:
            print(f"[{email}] No se pudo completar el campo de correo electrónico")
            return False
        fill_all_text_inputs_threaded(driver)
        select_random_options_in_all_dropdowns_threaded(driver)
        fill_all_radio_groups_threaded(driver)
        fill_all_checkboxes_threaded(driver)
        
        human_pause(0.1, 0.3)
        
        # Verificar ventana activa antes de avanzar
        try:
            _ = driver.title
        except Exception as e:
            print(f"[{email}] Ventana inactiva antes de avanzar: {e}")
            return False
            
        # Intentar avanzar o enviar
        if click_button_by_text_threaded("siguiente", driver):
            continue
        elif click_button_by_text_threaded("continuar", driver) or click_button_by_text_threaded("next", driver):
            continue
        else:
            # Intentar enviar
            print(f"[{email}] Preparando envío...")
            try:
                _ = driver.title
            except Exception as e:
                print(f"[{email}] Ventana inactiva al enviar: {e}")
                return False
            
            if attempt_submit_threaded(driver):
                print(f"[{email}] Respuesta enviada correctamente")
                return True
            else:
                print(f"[{email}] No se encontró un botón de envío disponible")
                return False

    print(f"[{email}] Límite de secciones alcanzado sin enviar.")
    return False

# Funciones thread-safe para llenar campos
def fill_email_field_threaded(driver, email, wait_obj=None):
    """Busca el campo de correo y lo completa; devuelve True si se llenó con éxito."""

    def _fill_input(inp):
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", inp)
            human_pause(0.05, 0.12)
            try:
                inp.clear()
            except Exception:
                pass
            inp.send_keys(email)
            human_pause(0.05, 0.12)
            return (inp.get_attribute('value') or '').strip() == email
        except Exception:
            return False

    try:
        if wait_obj is not None:
            try:
                wait_obj.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email'], input[type='text'], textarea")))
            except Exception:
                pass

        email_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='email'], input[type='text'], textarea")
        for inp in email_inputs:
            if not inp.is_displayed() or inp.get_attribute('disabled') or inp.get_attribute('readonly'):
                continue
            current = (inp.get_attribute('value') or '').strip()
            if current:
                # Campo ya tiene dato, se asume completado
                if '@' in current:
                    return True
                continue

            attr_bundle = " ".join([
                (inp.get_attribute('aria-label') or ''),
                (inp.get_attribute('name') or ''),
                (inp.get_attribute('placeholder') or ''),
                (inp.get_attribute('autocomplete') or ''),
            ]).lower()

            if any(key in attr_bundle for key in ('correo', 'email', 'e-mail', 'mail', 'correo electrónico', 'correo electronico')):
                if _fill_input(inp):
                    return True

        # Fallback directo por tipo email o nombres comunes
        fallback_selectors = [
            "input[type='email']",
            "input[name*='mail']",
            "input[name*='correo']",
            "input[autocomplete='email']",
        ]
        for sel in fallback_selectors:
            for inp in driver.find_elements(By.CSS_SELECTOR, sel):
                if not inp.is_displayed() or inp.get_attribute('disabled'):
                    continue
                if _fill_input(inp):
                    return True
    except Exception:
        pass

    return False

def fill_all_text_inputs_threaded(driver):
    """Rellena los campos de texto libres dentro del contexto seguro del hilo."""
    try:
        fields = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
    except Exception:
        return False
    
    for el in fields:
        try:
            if not el.is_displayed() or el.get_attribute('disabled') or el.get_attribute('readonly'):
                continue
            val = (el.get_attribute('value') or '').strip()
            if val:
                continue
                
            label = (el.get_attribute('aria-label') or '').strip().lower()
            if not label:
                try:
                    near = el.find_element(By.XPATH, "ancestor::*[1]//label|ancestor::*[2]//label|ancestor::*[2]//*[self::span or self::div][normalize-space()][1]")
                    label = (near.text or '').strip().lower()
                except Exception:
                    label = ''

            to_fill = None
            if any(k in label for k in ['correo', 'email', 'e-mail', 'mail']):
                to_fill = CURRENT_EMAIL
            elif any(k in label for k in ['nombre', 'name']):
                to_fill = _random_name()
            elif any(k in label for k in ['apellido', 'last name', 'apellidos']):
                to_fill = _random_name().split(' ')[1]
            elif any(k in label for k in ['dni', 'cédula', 'cedula', 'documento', 'id']):
                to_fill = ''.join(random.choice('0123456789') for _ in range(8))
            elif any(k in label for k in ['tel', 'cel', 'whatsapp', 'phone']):
                to_fill = _random_phone()
            elif any(k in label for k in ['edad', 'age']):
                to_fill = str(random.randint(18, 60))
            elif any(k in label for k in ['fecha', 'date']):
                to_fill = _random_date()
            elif any(k in label for k in ['hora', 'time']):
                to_fill = _random_time()
            else:
                to_fill = random.choice(['Sí', 'Ok', 'N/A', 'Acepto', _random_name()])

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
            human_pause(0.05, 0.12)
            try:
                el.clear()
            except Exception:
                pass
            el.send_keys(to_fill)
        except Exception:
            continue
    return True

def select_random_options_in_all_dropdowns_threaded(driver):
    """Selecciona opciones aleatorias en los desplegables visibles usando el driver del hilo."""
    try:
        dropdowns = driver.find_elements(By.CSS_SELECTOR, "div[role='listbox']")
        for dd in dropdowns:
            if not dd.is_displayed():
                continue
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dd)
                human_pause(0.05, 0.12)
                dd.click()
                human_pause(0.1, 0.2)
                options = driver.find_elements(By.CSS_SELECTOR, "div[role='option']")
                visible_opts = [o for o in options if o.is_displayed()]
                if visible_opts:
                    choice = random.choice(visible_opts)
                    choice.click()
                    human_pause(0.1, 0.2)
            except Exception:
                continue
    except Exception:
        pass

def fill_all_radio_groups_threaded(driver):
    """Marca opciones aleatorias en los grupos de radio disponibles para el hilo."""
    try:
        radio_groups = driver.find_elements(By.CSS_SELECTOR, "div[role='radiogroup']")
        for rg in radio_groups:
            if not rg.is_displayed():
                continue
            try:
                options = rg.find_elements(By.CSS_SELECTOR, "div[role='radio']")
                visible_opts = [o for o in options if o.is_displayed()]
                if visible_opts:
                    choice = random.choice(visible_opts)
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", choice)
                    human_pause(0.05, 0.12)
                    choice.click()
                    human_pause(0.1, 0.2)
            except Exception:
                continue
    except Exception:
        pass

def fill_all_checkboxes_threaded(driver):
    """Marca aleatoriamente casillas de verificación visibles en la instancia del hilo."""
    try:
        checkboxes = driver.find_elements(By.CSS_SELECTOR, "div[role='checkbox']")
        for cb in checkboxes:
            if not cb.is_displayed():
                continue
            if random.choice([True, False]):
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cb)
                    human_pause(0.05, 0.12)
                    cb.click()
                    human_pause(0.1, 0.2)
                except Exception:
                    continue
    except Exception:
        pass

def click_button_by_text_threaded(texto, driver):
    """Localiza y pulsa un botón por texto utilizando el WebDriver asignado al hilo."""
    texto = (texto or '').strip().lower()
    synonyms = {
        'next': ['siguiente', 'continuar', 'next', 'continue'],
        'submit': ['enviar', 'enviar respuestas', 'enviar respuesta', 'submit', 'send']
    }
    candidates = [texto]
    if texto in ('siguiente', 'next', 'continuar', 'continue'):
        candidates = synonyms['next']
    elif texto in ('enviar', 'submit', 'send', 'enviar respuesta', 'enviar respuestas'):
        candidates = synonyms['submit']

    try:
        driver.execute_script("window.scrollTo(0, 0);")
        human_pause(0.05, 0.12)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        human_pause(0.05, 0.12)
    except Exception:
        pass

    xpaths = []
    for t in candidates:
        xpaths.extend([
            f"//div[@role='button'][.//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')] or contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')]",
            f"//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')]",
            f"//a[@role='button' and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')]",
            f"//*[@role='button' and contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜÑ', 'abcdefghijklmnopqrstuvwxyzáéíóúüñ'), '{t}')]",
        ])

    found = []
    for xp in xpaths:
        try:
            elems = driver.find_elements(By.XPATH, xp)
            for el in elems:
                try:
                    if el.is_displayed():
                        rect = el.rect or {}
                        y = rect.get('y', 0)
                        label = (el.text or el.get_attribute('aria-label') or '').strip().lower()
                        score = 0
                        if 'enviar' in label:
                            score += 50
                        score += min(2000, y)
                        if any(x in label for x in ['atrás', 'anterior', 'borrar', 'limpiar', 'clear', 'back']):
                            score -= 1000
                        found.append((score, y, el))
                except Exception:
                    continue
        except Exception:
            continue

    if not found:
        return False

    found.sort(key=lambda t: (t[0], t[1]), reverse=True)

    for _, _, btn in found[:3]:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            human_pause(0.05, 0.12)
            btn.click()
            return True
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", btn)
                return True
            except Exception:
                continue

    return False

def attempt_submit_threaded(driver):
    """Aplica múltiples estrategias de envío dentro del flujo concurrente."""
    # 1) Por botones de texto conocidos
    if click_button_by_text_threaded("enviar", driver) or click_button_by_text_threaded("submit", driver):
        return True
    
    # 2) Intentar encontrar imagen enviar.png
    try:
        enviar_imgs = driver.find_elements(By.XPATH, "//img[contains(translate(@src, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'enviar')]")
        for img in enviar_imgs:
            if img.is_displayed():
                try:
                    target = img
                    try:
                        target = img.find_element(By.XPATH, "ancestor::*[@role='button' or self::button or self::a][1]")
                    except Exception:
                        pass
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
                    human_pause(0.05, 0.12)
                    target.click()
                    return True
                except Exception:
                    continue
    except Exception:
        pass
    
    # 3) Enter en el formulario como último recurso
    try:
        form = driver.find_element(By.CSS_SELECTOR, 'form')
        form.send_keys(Keys.ENTER)
        return True
    except Exception:
        pass
    
    return False

# === Flujo principal con threading ===

def main_with_threads():
    """Función principal que maneja múltiples hilos"""
    print(f"\n=== INICIANDO BOT CON {NUM_THREADS} HILOS ===")
    
    # Crear y lanzar hilos
    threads = []
    for i in range(NUM_THREADS):
        thread = threading.Thread(target=worker_thread, args=(i+1,), daemon=True)
        threads.append(thread)
        thread.start()
        print(f"Hilo {i+1} iniciado")
        
        # Pequeña pausa entre inicios de hilos para evitar problemas de concurrencia
        time.sleep(2)
    
    # Esperar a que todos los hilos terminen
    print(f"\nEsperando a que terminen todos los hilos...")
    for thread in threads:
        thread.join()
    
    # Esperar a que se complete la cola
    email_queue.join()
    
    print(f"\n=== PROCESO COMPLETADO ===")
    print(f"Envíos exitosos: {success_count}")
    print(f"Correos procesados: {len(emails_to_use) - email_queue.qsize()}")

# Ejecutar flujo principal
if __name__ == "__main__":
    opcion = menu_principal()
    generated_file = None
    if opcion == '2':
        generated_file = generate_emails_interactive()
    # Permitir que el usuario confirme si quiere usar el archivo recién generado u otro
    if opcion == '2':
        use_generated = input(f"¿Usar el archivo generado '{generated_file}'? (S/n): ").strip().lower()
        if use_generated in ('', 's', 'si', 'sí', 'y', 'yes'):
            default_file = generated_file
        else:
            default_file = None
    else:
        default_file = None

    FORM_URL_local, NUM_THREADS_local, MAX_SUBMISSIONS_local, EMAILS_FILE_local = get_user_input(default_emails_file=default_file)
    # Actualizar globales
    FORM_URL = FORM_URL_local
    NUM_THREADS = NUM_THREADS_local
    MAX_SUBMISSIONS = MAX_SUBMISSIONS_local
    EMAILS_FILE = EMAILS_FILE_local
    globals()['FORM_URL'] = FORM_URL
    globals()['NUM_THREADS'] = NUM_THREADS
    globals()['MAX_SUBMISSIONS'] = MAX_SUBMISSIONS
    globals()['EMAILS_FILE'] = EMAILS_FILE

    env_max = os.environ.get("MAX_SUBMISSIONS")
    if env_max and MAX_SUBMISSIONS is None:
        try:
            MAX_SUBMISSIONS = max(1, int(env_max))
            globals()['MAX_SUBMISSIONS'] = MAX_SUBMISSIONS
        except Exception:
            pass
    globals()['REUSE_SINGLE_BROWSER'] = _parse_bool_env(os.environ.get("REUSE_SINGLE_BROWSER"), REUSE_SINGLE_BROWSER)
    globals()['ALLOW_MANUAL_LOGIN'] = _parse_bool_env(os.environ.get("ALLOW_MANUAL_LOGIN"), ALLOW_MANUAL_LOGIN)

    print(f"\nConfiguración:")
    print(f"- Formulario: {FORM_URL}")
    print(f"- Hilos simultáneos: {NUM_THREADS}")
    print(f"- Límite de envíos: {MAX_SUBMISSIONS if MAX_SUBMISSIONS else 'Sin límite'}")
    print(f"- Archivo de correos: {EMAILS_FILE}")
    print("=" * 50)

    initialize_emails()
    main_with_threads()
