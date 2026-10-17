const CACHE='boq-nego-v4';
const SHELL=['./','./index.html','./manifest.json','./icon-192.png','./icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const req=e.request; if(req.method!=='GET')return;
  const url=new URL(req.url);
  if(url.origin===location.origin){
    // network first for the page so updates arrive; fall back to cache offline
    e.respondWith(fetch(req).then(res=>{const cp=res.clone();caches.open(CACHE).then(c=>c.put(req,cp));return res}).catch(()=>caches.match(req).then(r=>r||caches.match('./index.html'))));
  }else if(/fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)){
    e.respondWith(caches.match(req).then(hit=>hit||fetch(req).then(res=>{const cp=res.clone();caches.open(CACHE).then(c=>c.put(req,cp));return res})));
  }
});
