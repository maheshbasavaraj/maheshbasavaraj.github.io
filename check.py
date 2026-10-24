import subprocess,re,pathlib,pymupdf,sys
R=pathlib.Path(__file__).parent; C=r"C:\Program Files\Google\Chrome\Application\chrome.exe"
def dump(path,js,size):
    s=(R/path).read_text(encoding='utf8')
    pr="<script>window.addEventListener('load',()=>setTimeout(()=>{document.body.insertAdjacentHTML('beforeend','<pre id=probe>'+("+js+")+'</pre>')},2500))</script></body>"
    t=R/'probe.html'; t.write_text(s.replace('</body>',pr),encoding='utf8')
    o=subprocess.run([C,"--headless=new","--disable-gpu","--virtual-time-budget=8000","--window-size="+size,"--dump-dom",t.resolve().as_uri()],capture_output=True,text=True,encoding='utf8').stdout
    t.unlink(); return re.findall(r'<pre id="probe">(.*?)</pre>',o,re.S)
if sys.argv[1]=='resume':
    print(dump('resume.html',"[...document.querySelectorAll('section.page')].map((pg,k)=>{const P=pg.getBoundingClientRect();return [...pg.querySelectorAll('aside,main')].map((el,i)=>k+'.'+i+' pageH='+Math.round(P.height)+' bottom='+Math.round([...el.children].at(-1).getBoundingClientRect().bottom-P.top)).join('|')}).join('|')",'794,1123'))
    d=pymupdf.open(R/'Mahesh_Basavaraj_Resume.pdf');print('pages',len(d));print(d[1].get_text()[-200:])
else:
    print(dump('index.html',"'cards='+document.querySelectorAll('#proj a.card').length+' live='+document.querySelectorAll('#live .lr').length+' liveText='+document.getElementById('live').innerText.slice(0,200)+' navProjects='+!!document.querySelector('nav a[href=\"#projects\"]')+' ghLink='+!!document.querySelector('a[href=\"https://github.com/maheshbasavaraj\"]')",'1280,900'))
