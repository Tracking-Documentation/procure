# สร้าง book.html (หน้าจองบนเว็บเดียวกับระบบ) จาก gas/Booking.html — รันใหม่ทุกครั้งที่แก้ Booking.html:  python3 .mkbook.py
import re,os
d=os.path.dirname(os.path.abspath(__file__))
src=open(os.path.join(d,'gas/Booking.html'),encoding='utf-8').read()
idx=open(os.path.join(d,'index.html'),encoding='utf-8').read()
baked=re.search(r"SRV_BAKED='([^']+)'",idx).group(1)
shim='''<body>
<script>
/* ทำงานนอก Apps Script: อ่านโครงการ/โทเคนจากลิงก์ แล้วเรียกเซิร์ฟเวอร์ผ่าน fetch (ไม่ส่งคุกกี้ Google จึงไม่ติด /u/1/) */
(function(){var q=new URLSearchParams(location.search);document.body.setAttribute('data-p',q.get('p')||'');document.body.setAttribute('data-t',q.get('t')||'');
var SRV=q.get('s')||'%s';var OLD='กรุณาเข้าสู่ระบบ หรือใส่รหัสเจ้าหน้าที่';
function mk(ok,no){return new Proxy({},{get:function(_,k){if(k==='withSuccessHandler')return function(f){return mk(f,no)};if(k==='withFailureHandler')return function(f){return mk(ok,f)};
  return function(){var a=[].slice.call(arguments);fetch(SRV,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'book.'+k,args:a})}).then(function(r){return r.json()}).then(function(j){
    if(j.ok){if(ok)ok(j.result);return}
    if(j.error===OLD){location.replace(SRV+'?p='+encodeURIComponent(q.get('p')||'')+'&t='+encodeURIComponent(q.get('t')||''));return}
    if(no)no(new Error(j.error||'เกิดข้อผิดพลาด'))}).catch(function(){if(no)no(new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ กรุณาตรวจอินเทอร์เน็ตแล้วลองใหม่'))})}}})}
window.google={script:{run:mk(null,null)}}})();
</script>'''%baked
out=src.replace('<base target="_top">\n','')
out,n=re.subn(r'<body data-p="<\?= projectId \?>" data-t="<\?= token \?>">',lambda m:shim,out)
assert n==1 and '<?' not in out
out=out.replace('<head>','<head>\n<!-- ไฟล์นี้สร้างอัตโนมัติจาก gas/Booking.html ด้วย .mkbook.py — อย่าแก้ตรงนี้ -->\n<meta name="robots" content="noindex">\n<title>จองวันลงนามสัญญา</title>',1)
open(os.path.join(d,'book.html'),'w',encoding='utf-8').write(out)
print('book.html',len(out))
