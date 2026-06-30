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
    page.get_by_role("textbox", name="Contraseña Nueva Contraseña").click()
    page.get_by_role("textbox", name="Contraseña Nueva Contraseña").fill("***REDACTED*** (guardada cifrada en cred_adryan.bin)")
    page.get_by_role("button", name="INICIAR SESIÓN").click()
    page.get_by_role("link", name="Personal").click()
    page.get_by_role("link", name="Personal").click()
    page.get_by_role("link", name="Maestro del Personal sin").click()
    page.locator(".ctn-iconos-edit-tabla > a").first.click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
