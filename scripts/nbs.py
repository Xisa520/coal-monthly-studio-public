"""国家统计局官网数据发布栏目抓取：仅官网域名、限时、仅提取可证实字段，采纳需人工点击。"""
import re,datetime
from urllib.parse import urljoin,urlparse
import httpx
from bs4 import BeautifulSoup

BASE='https://www.stats.gov.cn/sj/zxfb/'
UA={'User-Agent':'CoalMarketMonthlyStudio/1.0 (research; local-user-initiated)'}

def clean(x):return re.sub(r'\s+','',x or '').replace('－','-').replace('—','—')

def get_html(url):
    p=urlparse(url)
    if p.scheme!='https' or p.hostname not in {'www.stats.gov.cn','stats.gov.cn'}:raise ValueError('只允许国家统计局官网HTTPS地址')
    with httpx.Client(timeout=12,follow_redirects=True,headers=UA,trust_env=True) as client:
        r=client.get(url);r.raise_for_status()
        if r.url.host not in {'www.stats.gov.cn','stats.gov.cn'}:raise ValueError('站点重定向到非官方域名')
        if len(r.content)>6_000_000:raise ValueError('网页超过内容限制')
        r.encoding=r.encoding or 'utf-8'
        return r.text

def discover():
    soup=BeautifulSoup(get_html(BASE),'html.parser')
    found={}
    for a in soup.select('a[href]'):
        name=clean(a.get_text(' ',strip=True))
        match=re.search(r'(20\d{2})年(\d{1,2})月份能源生产情况',name)
        if not match:continue
        url=urljoin(BASE,a.get('href'))
        if urlparse(url).hostname not in {'www.stats.gov.cn','stats.gov.cn'}:continue
        k=f'{match.group(1)}-{int(match.group(2)):02d}'
        found[k]={'month':k,'name':name,'url':url}
    return sorted(found.values(),key=lambda x:x['month'],reverse=True)

def read_energy(article):
    html=get_html(article['url'])
    soup=BeautifulSoup(html,'html.parser')
    # stats.gov.cn正文通常在含文档内容的div中；当前仅使用紧跟“一、原煤”的文本片段
    title=article['name']; body=clean(soup.get_text(' ',strip=True))
    body=body[body.find('一、原煤'): ] if '一、原煤' in body else body
    body=body.split('二、电力生产')[0] if '二、电力生产' in body else body[:15000]
    patterns=[
     (r'原煤产量([0-9]+(?:\.[0-9]+)?)(亿吨|万吨)[，,、]同比(增长|下降|减少)([0-9.]+)%','规上原煤产量','月度'),
     (r'原煤产量([0-9]+(?:\.[0-9]+)?)(亿吨|万吨)[，,、]同比(增长|下降|减少)([0-9.]+)%','规上原煤产量','累计')
    ]
    month=article['month']; out=[]
    # 按前缀分段确保月度与累计不混淆，不能靠第一个正则匹配猜测
    mlabel=f'{int(month[-2:])}月份'
    pos=body.find(mlabel)
    start=body[pos:pos+1100] if pos>=0 else body[:1100]
    cum=body.find(f'1—{int(month[-2:])}月份')
    if cum<0:cum=body.find(f'1-{int(month[-2:])}月份')
    cumtxt=body[cum:cum+750] if cum>=0 else ''
    regex=re.compile(r'(?:规上工业)?原煤产量(?:为|达)?\s*([\d.]+)(亿吨|万吨)[，,、]?(?:同比)?(增长|下降|减少)([\d.]+)%')
    for scope,chunk in [('月度',start),('累计',cumtxt)]:
        if not chunk:continue
        found=regex.search(chunk)
        if found:
            value=float(found.group(1)); unit=found.group(2)
            pct=float(found.group(4))*(-1 if found.group(3)!='增长' else 1)
            out.append({'report_month':month,'indicator':('规上原煤产量-'+scope),'value':value,'unit':unit,
              'year_on_year':pct,'url':article['url'],
              'release_date':None,'raw_excerpt':chunk[max(0,found.start()-60):found.end()+40]})
    return {'title':title,'url':article['url'],'month':month,'candidates':out,
            'note':'正则初筛候选值，可能因官方正文排版变化而缺失；须核对原文后点击“采纳”。'}

def fetch_latest():
    res=discover()
    if not res:raise ValueError('官网列表中未找到“能源生产情况”文章，可使用人工上传')
    return read_energy(res[0])
