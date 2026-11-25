from pathlib import Path
import http.server
import socketserver
import threading
import re
from playwright.sync_api import sync_playwright

ROOT=Path(__file__).resolve().parents[1]
class Quiet(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *args): pass

if __name__=='__main__':
    import os
    os.chdir(ROOT)
    srv=socketserver.TCPServer(('127.0.0.1',0),Quiet)
    thread=threading.Thread(target=srv.serve_forever,daemon=True);thread.start()
    link=(ROOT/'index.html').as_uri()
    with sync_playwright() as pw:
        browser=pw.chromium.launch(headless=True,executable_path='/usr/bin/chromium',args=['--no-sandbox'])
        page=browser.new_page(viewport={'width':1280,'height':860})
        errors=[]
        page.on('pageerror',lambda e: errors.append(str(e)))
        html=(ROOT/'index.html').read_text(encoding='utf8')
        html=re.sub(r'<script[^>]+src=\"[^\"]+\"[^>]*></script>', '', html)
        page.set_content(html)
        app=(ROOT/'app.js').read_text(encoding='utf8').replace("document.addEventListener('DOMContentLoaded',init);",'init();')
        page.evaluate(app)
        page.wait_for_selector('[data-section="dashboard"].active')
        assert page.locator('h1').first.inner_text().startswith('煤炭市场月报')
        page.locator('[data-view="report"]').click()
        page.locator('#reportAbstract').fill('本月市场走势待数据核实。')
        page.locator('#editor_production').fill('截至指定月份，原煤生产数据需以官方统计口径核对。')
        page.locator('#saveDraft').click()
        assert '保存' in page.locator('#toast').inner_text()
        page.locator('#qualityCheck').click()
        assert '尚未编写' in page.locator('#checkOutput').inner_text()
        page.locator('[data-view="ai"]').click()
        page.locator('#makePrompt').click()
        assert '禁止编造' in page.locator('#aiPrompt').input_value() or '不编造' in page.locator('#aiPrompt').input_value()
        page.locator('[data-view="official"]').click()
        page.wait_for_timeout(600)
        assert '候选' in page.locator('#officialList').inner_text() or '暂无' in page.locator('#officialList').inner_text() or '无法读取' in page.locator('#officialList').inner_text()
        page.locator('[data-view="upload"]').click()
        # stub only in test runtime; real browser uses pinned XLSX CDN
        page.evaluate('''() => {
          window.XLSX={
            read:()=>({SheetNames:['Sheet1'],Sheets:{Sheet1:{data:[['日期','煤价'],['2026-01-01',710],['2026-01-02',725]]}}}),
            utils:{sheet_to_json:(s)=>s.data,book_new:()=>({}),json_to_sheet:()=>({}),book_append_sheet:()=>{},writeFile:()=>{}},
            SSF:{parse_date_code:()=>null}
          };
        }''')
        page.locator('#fileInput').set_input_files({'name':'public_test.xlsx','mimeType':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','buffer':b'fake'})
        page.wait_for_function('document.getElementById("sheetSelect").options.length > 0'); assert 'Sheet1' in page.locator('#sheetSelect').inner_text()
        page.locator('#indicatorName').fill('煤价')
        page.locator('#indicatorUnit').fill('元/吨')
        page.locator('#importSeries').click()
        assert page.locator('#metricRows').inner_text()=='2'
        page.locator('#chartIndicator').select_option(label='煤价（元/吨）')
        page.locator('[data-view="dashboard"]').click()
        assert '725' in page.locator('#chartStats').inner_text()
        # Word: mock the zip helper to inspect generated OpenXML payloads
        page.evaluate('''() => {
          window.JSZip=class{
            constructor(){this.files={};}
            file(n,d){this.files[n]=d;}
            generateAsync(){window.__generatedFiles=this.files; return Promise.resolve(new Blob(['mock']));}
          };
        }''')
        page.locator('[data-view="report"]').click()
        page.locator('#downloadDocx').click()
        page.wait_for_function('!!window.__generatedFiles')
        xmls=page.evaluate('''() => ({
          doc:window.__generatedFiles['word/document.xml'],
          styles:window.__generatedFiles['word/styles.xml'],
          types:window.__generatedFiles['[Content_Types].xml'],
          rels:window.__generatedFiles['_rels/.rels'],
          docrels:window.__generatedFiles['word/_rels/document.xml.rels'],
          image:window.__generatedFiles['word/media/chart.png']
        })''')
        import xml.etree.ElementTree as ET
        for item in [xmls[k] for k in ('doc','styles','types','rels','docrels')]: ET.fromstring(item)
        assert '煤炭市场' in xmls['doc']
        import zipfile,io
        from docx import Document
        output=io.BytesIO()
        with zipfile.ZipFile(output,'w') as z:
          for key,name in [('doc','word/document.xml'),('styles','word/styles.xml'),('types','[Content_Types].xml'),('rels','_rels/.rels'),('docrels','word/_rels/document.xml.rels')]:
            z.writestr(name,xmls[key])
          if xmls.get('image'):
            import base64
            z.writestr('word/media/chart.png',base64.b64decode(xmls['image']))
        output.seek(0)
        document=Document(output)
        assert document.paragraphs[0].text.startswith('煤炭市场') or '煤炭市场' in document.paragraphs[0].text
        page.screenshot(path=str(ROOT/'tests'/'ui-check.png'),full_page=True)
        assert not errors,errors
        browser.close()
    srv.shutdown()
    print('PASS: navigation, local draft, QA, AI prompt, official JSON, local file import, chart and no JS errors')
