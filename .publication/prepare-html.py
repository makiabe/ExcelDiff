from bs4 import BeautifulSoup, NavigableString
from pathlib import Path
import re,json
root=Path.cwd()
soup=BeautifulSoup((root/'.publication/index-original.html').read_text(),'html.parser')
ja=json.loads((root/'locales/ja.json').read_text())
counter=0
for el in list(soup.find_all(string=True)):
 if el.parent.name in ['script','style']: continue
 text=str(el)
 if re.search(r'[\u3040-\u30ff\u3400-\u9fff]',text):
  counter+=1;key=f'ui.{counter:03d}';ja[key]=text.strip();el.replace_with(NavigableString('{{'+key+'}}'))
for tag in soup.find_all():
 for attr in ['alt','title','placeholder','aria-label','content']:
  if tag.has_attr(attr) and re.search(r'[\u3040-\u30ff\u3400-\u9fff]',tag[attr]):
   counter+=1;key=f'ui.{counter:03d}';ja[key]=tag[attr];tag[attr]='{{'+key+'}}'
assert counter==94,'Unexpected source HTML; stop instead of translating the wrong text.'
soup.html['lang']='{{lang}}';soup.html['dir']='{{dir}}'
soup.title.string='{{seo.title}}'
soup.find('meta',attrs={'name':'description'})['content']='{{seo.description}}'
for key in ['canonical','alternates','social','structured','stylesheet']:
 soup.head.append(NavigableString('{{raw.'+key+'}}'))
for tag in soup.find_all('script',src=True):
 if not tag['src'].startswith(('https://','http://')): tag['src']='{{base}}'+tag['src'].split('?')[0]+'?v=1.2.0'
soup.find('link',rel='icon')['href']='{{base}}favicon.ico?v=frog-cat-1'
nav=soup.header.find('nav');nav['class']=['primary-nav']
soup.header.div.append(NavigableString('{{raw.languageSwitcher}}'))
soup.header.div['class']=['header-inner']
hero=soup.main.find('section');hero['class']=['hero-section']
hero.find('div',recursive=False)['class']=['hero-copy']
hero_img=soup.find('img',src='assets/hero-cat.webp');hero_img['src']='{{asset.hero}}';hero_img['alt']='{{mascot.heroAlt}}'
hero_img['width']='{{asset.heroWidth}}';hero_img['height']='{{asset.heroHeight}}';hero_img['fetchpriority']='high'
hero_img.parent['class']=['mascot','hero-mascot'];hero_img.parent.append(NavigableString('{{raw.heroBubble}}'))
bottom=soup.find('img',src='assets/bottom-cat.webp');bottom['src']='{{asset.bottom}}';bottom['alt']='{{mascot.bottomAlt}}'
bottom['width']='{{asset.bottomWidth}}';bottom['height']='{{asset.bottomHeight}}'
bottom['loading']='lazy';bottom['style']='display:block;width:100%;height:auto;'
wrap=soup.new_tag('div');wrap['class']=['mascot','bottom-mascot'];bottom.wrap(wrap);wrap.append(NavigableString('{{raw.bottomBubble}}'))
soup.find(id='dropZone')['class']=['drop-zone']
options=soup.find(id='ignoreWhitespace').parent.parent.parent;options['class']=['options-grid']
for div in options.find_all('div',recursive=False):div['class']=['options-column']
soup.find(id='excludeColumns')['aria-label']='{{form.excludeLabel}}'
soup.find(id='primaryKey')['aria-label']='{{form.keyLabel}}'
soup.find(id='appVersion').string='v1.2.0'
first=soup.find('script',src=True)
for path in ['{{base}}locales/{{lang}}.js?v=1.2.0','{{base}}i18n.js?v=1.2.0']:
 s=soup.new_tag('script',src=path);first.insert_before(s)
no=soup.new_tag('noscript');no.string='{{ui.noScript}}';soup.main.insert(0,no)
ja.update({
 'seo.title':'Excel Diff - Excel / CSV 差分比較ツール【無料・アップロード不要】',
 'seo.description':'2つのExcel・CSVファイルをブラウザ内で比較。セル・数式・書式・行の差分を検出し、Excel・HTML・PDFレポートに出力。無料、登録・アップロード不要。',
 'seo.applicationDescription':'Excel・CSVファイルの差分をブラウザ内で比較する無料ツール。',
 'ui.noScript':'ファイルの比較にはJavaScriptを有効にしてください。説明とFAQはこのままお読みいただけます。',
 'nav.language':'言語を切り替える',
 'mascot.heroAlt':'カエルの帽子をかぶった猫とExcelファイル',
 'mascot.bottomAlt':'カエルの帽子をかぶった猫',
 'mascot.heroBubble':'違いを一緒に\n見つけよう！',
 'mascot.bottomBubble':'ファイルは\nあなたの端末の中に。',
 'form.excludeLabel':'比較対象外の列',
 'form.keyLabel':'主キー（行の一致判定に使用）'
})
ja['app.073']=ja['app.073'].replace('lang="ja"','lang="{locale}"')
(root/'src/index.template.html').write_text(str(soup))
(root/'locales/ja.json').write_text(json.dumps(ja,ensure_ascii=False,indent=2)+'\n')
