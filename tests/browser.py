"""Smoke-test a served Japanese/English site, including uploaded-file privacy."""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

base = (sys.argv[1] if len(sys.argv) > 1 else 'http://127.0.0.1:8765/').rstrip('/') + '/'
out = Path('test-results'); out.mkdir(exist_ok=True)
errors = []
with sync_playwright() as p:
    browser = p.chromium.launch()
    for width in [1440, 768, 375, 320]:
        page = browser.new_page(viewport={'width': width, 'height': 1000})
        page.on('pageerror', lambda e: errors.append(str(e)))
        response = page.goto(base + 'en/', wait_until='networkidle')
        assert response.status == 200, response.status
        assert page.locator('html').get_attribute('lang') == 'en'
        assert 'Compare Excel' in page.title()
        page.locator('img').evaluate_all('(xs) => xs.forEach(x => x.loading = "eager")')
        page.wait_for_function('Array.from(document.images).every(x => x.complete && x.naturalWidth > 0)')
        assert page.evaluate('document.documentElement.scrollWidth <= innerWidth'), width
        assert page.locator('.bubble').count() == 2
        assert 'No file selected' in page.locator('#fileChips').inner_text()
        page.screenshot(path=str(out / f'english-{width}.png'), full_page=True)
        if width == 1440:
            requests = []
            page.on('request', lambda r: requests.append(r.url))
            page.locator('#fileInput').set_input_files([
                {'name': 'before.csv', 'mimeType': 'text/csv', 'buffer': b'ID,Name\n1,Alpha\n2,Beta\n'},
                {'name': 'after.csv', 'mimeType': 'text/csv', 'buffer': b'ID,Name\n1,ALPHA\n3,Gamma\n'}
            ])
            assert 'before.csv' in page.locator('#fileChips').inner_text()
            assert 'after.csv' in page.locator('#fileChips').inner_text()
            page.locator('#primaryKey').fill('A')
            page.locator('#compareBtn').click()
            page.locator('#results').wait_for(state='visible')
            summary = page.locator('#resultSummary').inner_text()
            print('English comparison:', summary)
            for text in ['1 changed cells', '1 added rows', '1 deleted rows']:
                assert text in summary, summary
            assert not requests, f'Unexpected network activity while comparing CSV: {requests}'
            with page.expect_download() as download:
                page.locator('#exportHtml').click()
            report = out / 'comparison-report.html'
            download.value.save_as(str(report))
            assert '<html lang="en">' in report.read_text()
            with page.expect_download() as download:
                page.locator('#exportExcel').click()
            download.value.save_as(str(out / 'comparison-report.xlsx'))
            assert (out / 'comparison-report.xlsx').read_bytes()[:2] == b'PK'
            result = page.evaluate('''async () => {
                const blob = ExcelDiffCodec.writeReport([{name:'Example', rows:[['ID','Value'], ['001','<script>alert(1)</script>']]}]);
                const book = await ExcelDiffCodec.read(new File([blob], 'sample.xlsx'));
                return book.sheets[0].rows.get(2).get(0).v;
            }''')
            assert result == '001'
            page.screenshot(path=str(out / 'english-results.png'), full_page=True)
        page.close()
    page = browser.new_page(viewport={'width': 1440, 'height': 1000})
    page.on('pageerror', lambda e: errors.append(str(e)))
    response = page.goto(base, wait_until='networkidle')
    assert response.status == 200
    assert page.locator('html').get_attribute('lang') == 'ja'
    assert '未選択' in page.locator('#fileChips').inner_text()
    assert page.locator('a[hreflang="en"]').count() == 1
    page.locator('a[hreflang="en"]').click()
    page.wait_for_url(base + 'en/')
    assert page.locator('html').get_attribute('lang') == 'en'
    page.locator('a[hreflang="ja"]').click()
    page.wait_for_url(base)
    assert page.locator('html').get_attribute('lang') == 'ja'
    browser.close()
assert not errors, errors
print('Browser smoke tests passed, including both languages, responsive layouts, exports and local-only comparison.')
