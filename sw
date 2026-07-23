self.addEventListener('push', e=>{
  let d={}; try{ d=e.data.json(); }catch(_){ d={title:'새 문의', body:(e.data&&e.data.text())||''}; }
  e.waitUntil(self.registration.showNotification(d.title||'새 문의', {
    body:d.body||'', icon:'/icon-192.png', badge:'/icon-192.png', tag:d.tag||'staybut-chat',
    renotify:true, data:{url:d.url||'/console.html'}
  }));
});
self.addEventListener('notificationclick', e=>{
  e.notification.close();
  const url=(e.notification.data&&e.notification.data.url)||'/console.html';
  e.waitUntil(clients.matchAll({type:'window',includeUncontrolled:true}).then(cl=>{
    for(const c of cl){ if(c.url.includes('console')&&'focus'in c) return c.focus(); }
    return clients.openWindow(url);
  }));
});
