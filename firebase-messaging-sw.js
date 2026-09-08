const REPORTIA_SW_VERSION='53.3';
importScripts('./firebase-config.js');
importScripts('https://www.gstatic.com/firebasejs/12.18.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/12.18.0/firebase-messaging-compat.js');

const cfg=self.REPORTIA_FIREBASE_CONFIG||{};
const ready=cfg.apiKey && !String(cfg.apiKey).includes('PEGA_AQUI');
if(ready){
  firebase.initializeApp(cfg);
  const messaging=firebase.messaging();
  messaging.onBackgroundMessage(payload=>{
    const d=payload.data||{};
    self.registration.showNotification(d.title||'REPORT.IA',{body:d.body||'Tienes una nueva notificación.',icon:'./icon-192.png',badge:'./icon-192.png',tag:d.messageId||'reportia',renotify:true,data:{url:d.url||'./?reportiaNotification=1',module:d.module||'',messageId:d.messageId||''}});
  });
}
self.addEventListener('notificationclick',event=>{event.notification.close();const url=new URL(event.notification.data?.url||'./?reportiaNotification=1',self.registration.scope).href;event.waitUntil(clients.matchAll({type:'window',includeUncontrolled:true}).then(list=>{for(const c of list){if('focus'in c){c.navigate(url);return c.focus()}}return clients.openWindow?clients.openWindow(url):null}))});
self.addEventListener('fetch',()=>{});

self.addEventListener('install',event=>{self.skipWaiting();});
self.addEventListener('activate',event=>{event.waitUntil(self.clients.claim());});
