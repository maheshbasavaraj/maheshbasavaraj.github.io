import sys, pathlib, subprocess, re
R = pathlib.Path(__file__).parent
C = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
w = sys.argv[1] if len(sys.argv) > 1 else "390"
js = """(()=>{const o=[innerWidth+'vw docScroll='+document.documentElement.scrollWidth];
document.querySelectorAll('header,header>div,.side,.facts,.stats,nav .wrap,.lede,.cta,h1').forEach(e=>o.push((e.className||e.tagName)+':'+Math.round(e.getBoundingClientRect().right)));
return o.join(' | ')})()"""
s = (R / "index.html").read_text(encoding="utf8")
s = s.replace("</body>", "<script>window.addEventListener('load',()=>setTimeout(()=>{const p=document.createElement('pre');p.id='probe';p.textContent=" + js + ";document.body.appendChild(p)},1500))</script></body>")
t = R / "probe.html"; t.write_text(s, encoding="utf8")
o = subprocess.run([C, "--headless=new", "--disable-gpu", "--virtual-time-budget=6000", f"--window-size={w},1400", "--dump-dom", t.resolve().as_uri()], capture_output=True, text=True, encoding="utf8").stdout
t.unlink(); print(re.findall(r'<pre id="probe">(.*?)</pre>', o, re.S))
