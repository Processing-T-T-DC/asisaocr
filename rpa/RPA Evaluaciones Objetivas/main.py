import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from datetime import datetime

from constants import DOWNLOAD_DIR
from login import login
from eactivity import go_to_eactivity_and_act
from generar_informe import generar_informe
from excel_runner import (
    cargar_excel,
    iterar_filas,
    escribir_resultado,
    guardar_excel,
    obtener_columnas_resultado
)

sys.stdout.reconfigure(encoding="utf-8")

# =====================
# INIT DRIVER
# =====================
def init_driver():
    options = Options()
    options.add_argument("--start-maximized")

    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
    )

    return webdriver.Chrome(options=options)


def main():
    # =====================
    # CARGAR EXCEL
    # =====================
    wb, ws = cargar_excel()
    cols = obtener_columnas_resultado(ws)

    # =====================
    # INICIALIZAR SELENIUM
    # =====================
    driver = init_driver()

    try:
        # =====================
        # LOGIN
        # =====================
        login(driver)

        # =====================
        # PROCESAR FILAS
        # =====================
        for row, tratamiento, finalidad, nombre, descripcion in iterar_filas(ws):
            print(f"\n▶ Fila {row} | {tratamiento} | {finalidad}")

            try:
                # =====================
                # EACTIVITY
                # =====================
                go_to_eactivity_and_act(
                    driver,
                    tratamiento=tratamiento,
                    finalidad=finalidad,
                    nombre=nombre,
                    descripcion=descripcion
                )

                # =====================
                # GENERAR INFORME
                # =====================
                file_name = generar_informe(
                    driver,
                    tratamiento,
                    finalidad,
                    nombre,
                    descripcion
                )

                escribir_resultado(ws, row, cols["estado"], "OK")
                escribir_resultado(ws, row, cols["fichero"], file_name)
                escribir_resultado(
                    ws,
                    row,
                    cols["fecha"],
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                )

            except Exception as e:
                escribir_resultado(ws, row, cols["estado"], "ERROR")
                escribir_resultado(ws, row, cols["error"], str(e))
                escribir_resultado(
                    ws,
                    row,
                    cols["fecha"],
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                )

    finally:
        guardar_excel(wb)
        driver.quit()

    print("✅ Proceso completo finalizado")


if __name__ == "__main__":
    main()
