"""Необязательная проверка интерфейса: требуется playwright и установленный Edge."""
from pathlib import Path
import json
import shutil
import tempfile
import threading

from playwright.sync_api import sync_playwright, expect

from test_webapp import new_plan
from webapp import Application, Handler, ThreadingHTTPServer

ROOT = Path(__file__).resolve().parent


def main():
    artifacts = ROOT / '.kini-app' / 'preview'
    artifacts.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        config = json.loads((ROOT / 'kini_config.json').read_text(encoding='utf-8'))
        config['template'] = str(ROOT / config['template'])
        for group in config['groups']:
            group['file'] = str(ROOT / group['file'])
        (root / 'kini_config.json').write_text(json.dumps(config), encoding='utf-8')
        shutil.copytree(ROOT / 'web', root / 'web')
        app = Application(root)
        server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        server.app = app
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(channel='msedge', headless=True)
                page = browser.new_page(viewport={'width':1440, 'height':1050}, device_scale_factor=1)
                errors = []
                page.on('pageerror', lambda e: errors.append(str(e)))
                page.goto(f'http://127.0.0.1:{server.server_port}')
                page.locator('#loading').wait_for(state='hidden')
                assert page.locator('.group').count() == 5
                assert page.locator('#metric-students').inner_text() == '35'
                page.screenshot(path=str(artifacts / 'desktop.png'), full_page=True)
                # Новый МУП и обязательная численность: без неё расчёт не запускается.
                path = root / 'Новый МУП CS 2031.xlsx'
                new_plan(path, shift=2)
                page.locator('#file-picker').set_input_files(str(path))
                expect(page.locator('.group')).to_have_count(6)
                group = page.locator('.group').last
                assert group.locator('[name=program]').input_value() == 'CS'
                assert group.locator('[name=intake]').input_value() == '2031'
                page.locator('#generate').click()
                assert group.locator('[name=students]').evaluate('(el) => !el.validity.valid')
                assert not app.jobs
                group.locator('[name=students]').fill('11')
                group.locator('[name=name]').fill('CS-31')
                # Оставляем новый набор, проверяем полный расчёт в браузере.
                for _ in range(5):
                    page.locator('.group').first.locator('.remove').click()
                page.locator('#generate').click()
                page.locator('#download-all').wait_for(state='visible', timeout=60000)
                assert page.locator('.result-card').count() == 3
                assert page.locator('.result-card').first.locator('h3').inner_text() == '2031–2032'
                with page.expect_download() as download:
                    page.locator('#download-all').click()
                assert download.value.suggested_filename == 'KINI_расчеты.zip'
                download.value.save_as(str(root / 'download.zip'))
                # Возвращаем пять исходных групп для контрольной выгрузки и снимка результатов.
                page.evaluate('(groups) => localStorage.setItem("kini-selection-v1", JSON.stringify(groups))', app.defaults)
                page.reload()
                page.locator('#loading').wait_for(state='hidden')
                assert page.locator('.group').count() == 5
                page.locator('#generate').click()
                expect(page.locator('.result-card')).to_have_count(4, timeout=60000)
                page.locator('#job-progress').wait_for(state='hidden')
                assert page.locator('#error').is_hidden()
                assert page.locator('#stale').is_hidden()
                page.locator('#issue-search').fill('математика')
                assert page.locator('#issue-counter').inner_text()
                page.locator('#issue-search').fill('')
                page.locator('#results').scroll_into_view_if_needed()
                page.screenshot(path=str(artifacts / 'results.png'))
                output = ROOT / 'output_browser'
                output.mkdir(exist_ok=True)
                for link in page.locator('.result-card a').all():
                    with page.expect_download() as download:
                        link.click()
                    download.value.save_as(str(output / download.value.suggested_filename))
                with page.expect_download() as download:
                    page.locator('#download-report').click()
                download.value.save_as(str(output / download.value.suggested_filename))
                page.set_viewport_size({'width':390, 'height':844})
                page.evaluate('scrollTo(0,0)')
                assert page.evaluate('document.documentElement.scrollWidth <= innerWidth'), 'Mobile horizontal overflow'
                page.screenshot(path=str(artifacts / 'mobile.png'))
                assert not errors, errors
                browser.close()
                print('Browser OK: upload, validation, two generations, download, filters, desktop/mobile, no JS errors.')
        finally:
            server.shutdown()
            server.server_close()
            thread.join()
            app.executor.shutdown()


if __name__ == '__main__':
    main()
