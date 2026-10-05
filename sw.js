// Super PDF offline cache
const CACHE='superpdf-7a1c420d45';
const FILES=["./", "./Sortable.min.js", "./UTIF.js", "./app.css", "./app.js", "./cmaps.json", "./fonts.json", "./icon-180.png", "./icon-192.png", "./icon-512.png", "./icon.svg", "./index.html", "./manifest.webmanifest", "./pako.min.js", "./pdf-lib.min.js", "./pdf.min.mjs", "./pdf.worker.min.mjs", "./scanic.umd.js", "./tiffenc.js"];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES.map(u=>new Request(u,{cache:'reload'})))).then(()=>self.skipWaiting()));});
const OCR='superpdf-ocr-1';
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE&&k!==OCR).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{if(e.request.method!=='GET')return;
  const u=new URL(e.request.url);
  if(u.pathname.includes('/ocr/')){ // Thai OCR engine (~16 MB): fetched on first use, then kept offline across app updates
    e.respondWith(caches.open(OCR).then(c=>c.match(e.request,{ignoreSearch:true}).then(r=>r||fetch(e.request).then(res=>{if(res.ok)c.put(e.request,res.clone());return res;}))));return;}
  e.respondWith(caches.match(e.request,{ignoreSearch:true}).then(r=>r||fetch(e.request)));});
