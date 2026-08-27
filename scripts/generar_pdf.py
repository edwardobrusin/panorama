from playwright.sync_api import sync_playwright
import os


def generar_pdf_panorama(url: str = "http://localhost:8501",
                          output_path: str = "panorama_economico.pdf") -> None:
    """
    Genera el PDF del dashboard 'Panorama Económico' (dash_v2.py) en tamaño
    Carta, paginado de forma nativa por el motor de impresión de Chromium.

    A diferencia del motor anterior (una sola hoja de altura variable), aquí
    NO forzamos una altura custom: dejamos que Chromium reparta el contenido
    en tantas hojas Carta como haga falta, respetando las reglas CSS
    'break-inside: avoid' ya definidas en dash_v2.py (cada .metric-container,
    cada .cat-bloque y cada <tr> son unidades indivisibles).
    """
    with sync_playwright() as p:
        print("🤖 Iniciando navegador Chromium (Headless)...")
        browser = p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-dev-shm-usage']
        )
        context = browser.new_context(
            viewport={"width": 1300, "height": 1080},
            device_scale_factor=2,  # texto y bordes de tabla más nítidos en el PDF
        )
        page = context.new_page()

        print(f"🌐 Cargando {url} ...")
        page.goto(url, wait_until="domcontentloaded")

        # Esperamos a que el contenido principal esté renderizado
        page.wait_for_selector(".main-header", state="visible", timeout=60000)
        page.wait_for_selector(".metric-container", state="visible", timeout=60000)
        page.wait_for_load_state("networkidle", timeout=60000)

        print("🧹 Inyectando CSS de limpieza para exportación...")
        css_export = """
            /* Ocultar chrome de Streamlit que no debe aparecer en el PDF */
            [data-testid="stSidebar"], header[data-testid="stHeader"], footer,
            .stAppDeployButton, #MainMenu, .floating-download-btn, .floating-warning,
            [data-testid="stDecoration"], [data-testid="stToolbar"],
            [data-testid="stStatusWidget"] {
                display: none !important;
            }
            html, body, .stApp, [data-testid="stAppViewContainer"],
            [data-testid="stMain"], .main {
                height: auto !important;
                min-height: 0 !important;
                overflow: visible !important;
                position: static !important;
                margin-top: 0 !important;
                padding-top: 0 !important;
            }
            /* Padding separado por lado: arriba lo llevamos casi a cero (solo el
               PDF puede darse ese lujo, ya que aquí no reservamos espacio para
               ninguna toolbar), dejando aire normal en los otros tres lados. */
            /* Especificidad reforzada a propósito (prefijo "body") para NO depender
               de qué <style> se haya inyectado al DOM más recientemente. Así gana
               siempre, incluso si Streamlit vuelve a insertar su propio <style>
               después de un rerun disparado por los clics en los banners. */
            body [data-testid="stAppViewBlockContainer"],
            body .block-container,
            body [data-testid="block-container"] {
                max-width: 100% !important;
                padding-top: 0.05in !important;
                padding-right: 0.1in !important;
                padding-bottom: 0.3in !important;
                padding-left: 0.1in !important;
                margin: 0 !important;
            }
            .stApp { background-color: #F8FAFC !important; }
        """

        def inyectar_css_export():
            page.add_style_tag(content=css_export)

        inyectar_css_export()

        print("🔇 Cerrando nudges/banners nativos de Streamlit (si aparecen)...")

        # Opción A (preferida): puede haber MÁS DE UN banner simultáneo (p. ej.
        # "Help agents write better apps..." Y "Install the official Streamlit
        # skills..." al mismo tiempo, cada uno con su propio botón "Don't show
        # again"). Hacemos clic en todos los que existan, re-consultando el DOM
        # en cada vuelta porque tras cada clic el banner desaparece y las
        # referencias anteriores quedan obsoletas.
        intentos_click = 0
        while intentos_click < 5:
            botones = page.get_by_role("button", name="Don't show again", exact=False)
            total = botones.count()
            if total == 0:
                break
            print(f"      🔘 Encontrado(s) {total} botón(es) 'Don't show again', haciendo clic...")
            try:
                botones.first.click(timeout=3000)
                page.wait_for_timeout(300)  # deja que el DOM se asiente tras el clic
            except Exception:
                break
            intentos_click += 1

        if intentos_click == 0:
            print("      ⚠️ No se encontró ningún botón 'Don't show again' a tiempo.")

        # Opción B (respaldo): ocultar por texto cualquier tarjeta/nudge flotante que
        # haya quedado, sin depender de un data-testid interno de Streamlit que puede
        # cambiar entre versiones. Cubre ambos banners conocidos hasta ahora.
        page.evaluate("""
            () => {
                const FRASES_A_OCULTAR = [
                    "help agents write better apps",
                    "install the official streamlit skills",
                    "ai coding agents can build and debug your apps",
                ];
                const candidatos = document.querySelectorAll("body *");
                candidatos.forEach(el => {
                    if (el.children.length > 0) return;
                    const texto = (el.innerText || el.textContent || "").trim().toLowerCase();
                    if (!texto) return;
                    if (FRASES_A_OCULTAR.some(f => texto.includes(f))) {
                        const contenedor = el.closest('[data-testid], div, section') || el;
                        contenedor.style.setProperty("display", "none", "important");
                    }
                });
            }
        """)

        # Re-inyectamos el CSS de exportación justo antes de imprimir: si algún
        # rerun de Streamlit (por ejemplo, tras cerrar los banners) reescribió el
        # DOM y re-insertó su propio <style>, esto garantiza que nuestra versión
        # sea la última palabra de todos modos. Es redundante en el caso normal,
        # pero es la red de seguridad que evita el comportamiento intermitente.
        inyectar_css_export()
        page.wait_for_timeout(500)  # reflow del DOM tras cerrar nudges + CSS

        print(f"📄 Generando PDF paginado (Carta) en: {output_path}")
        page.pdf(
            path=output_path,
            format="Letter",
            print_background=True,
            margin={"top": "0in", "bottom": "0in", "left": "0in", "right": "0in"},
            display_header_footer=False,
        )

        browser.close()
        print(f"✅ ¡Listo! PDF generado en: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    generar_pdf_panorama(url="http://localhost:8501")