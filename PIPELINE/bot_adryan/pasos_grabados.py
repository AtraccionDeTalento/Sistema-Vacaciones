import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(channel="chrome", headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://adryancloudusil.sapia.com.pe/")
    page.get_by_role("textbox", name="Usuario Usuario").click()
    page.get_by_role("textbox", name="Usuario Usuario").fill("JLOPEZL4")
    page.get_by_role("textbox", name="Contraseña Nueva Contraseña").click()
    page.get_by_role("textbox", name="Contraseña Nueva Contraseña").fill("34frDA@#123")
    page.get_by_role("button", name="INICIAR SESIÓN").click()
    page.get_by_role("link", description="Personal", exact=True).click()
    page.get_by_role("link", name="Vacaciones por Ejercicio").click()
    page.get_by_role("link", description="Vacaciones", exact=True).click()
    page.locator("div").filter(has_text="Personal").nth(2).click()
    page.get_by_role("link", name="Vacaciones por Motivo").click()
    page.get_by_role("textbox", name="Fecha Inicio").click()
    page.get_by_role("combobox").nth(2).select_option("3")
    page.get_by_role("gridcell", name="01/04/").click()
    page.get_by_role("textbox", name="Fecha Término").click()
    page.get_by_role("combobox").nth(2).select_option("7")
    page.get_by_role("gridcell", name="31/08/").click()
    page.get_by_role("button", name="Buscar").click()
    with page.expect_download() as download_info:
        page.locator(".mr-3").click()
    download = download_info.value

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
