// Super PDF offline cache
const CACHE='superpdf-a27fd36ab9';
const FILES=["./", "./Sortable.min.js", "./UTIF.js", "./app.css", "./app.js", "./cmaps.json", "./fonts.json", "./icon-180.png", "./icon-192.png", "./icon-512.png", "./icon.svg", "./index.html", "./manifest.webmanifest", "./pako.min.js", "./pdf-lib.min.js", "./pdf.min.mjs", "./pdf.worker.min.mjs", "./tiffenc.js"];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES.map(u=>new Request(u,{cache:'reload'})))).then(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{if(e.request.method!=='GET')return;
  e.respondWith(caches.match(e.request,{ignoreSearch:true}).then(r=>r||fetch(e.request)));});
