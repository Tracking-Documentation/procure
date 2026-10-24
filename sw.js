const CACHE='boq-nego-v139';
const SHELL=['./','./index.html','./tor-tpl.js','./contract-tpl.js','./doh-seal.png','./sign-1.png','./sign-2.png','./wplan-1.jpg','./wplan-2.jpg','./wplan-3.jpg','./wplan-4.jpg','./manifest.json','./icon-r192.png','./icon-r512.png','./apple-touch-icon.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL.map(u=>new Request(u,{cache:'reload'})))).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const req=e.request; if(req.method!=='GET')return;
  const url=new URL(req.url);
  if(url.origin===location.origin){
    // always check the server for the newest version (bypass browser HTTP cache); use saved copy only when offline
    e.respondWith(fetch(req,{cache:'no-cache'}).then(res=>{const cp=res.clone();caches.open(CACHE).then(c=>c.put(req,cp));return res}).catch(()=>caches.match(req).then(r=>r||caches.match('./index.html'))));
  }else if(/fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)){
    e.respondWith(caches.match(req).then(hit=>hit||fetch(req).then(res=>{const cp=res.clone();caches.open(CACHE).then(c=>c.put(req,cp));return res})));
  }
});
