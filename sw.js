/* Service worker de SHAKA · TAREAS
   Solo se ocupa de las notificaciones: recibe el aviso y lo muestra,
   y al tocarlo abre la app. No cachea nada, para no servir versiones viejas. */

self.addEventListener('install', function(e){ self.skipWaiting(); });
self.addEventListener('activate', function(e){ e.waitUntil(self.clients.claim()); });

self.addEventListener('push', function(e){
  var d = {};
  try { d = e.data ? e.data.json() : {}; } catch (err) { d = { body: e.data && e.data.text() }; }
  var titulo = d.title || 'Shaka · Tareas';
  var opciones = {
    body: d.body || '',
    icon: 'icon-192.png',
    badge: 'icon-192.png',
    tag: d.tag || 'shaka-tarea',
    data: { url: d.url || './' },
    requireInteraction: false
  };
  e.waitUntil(self.registration.showNotification(titulo, opciones));
});

self.addEventListener('notificationclick', function(e){
  e.notification.close();
  var destino = (e.notification.data && e.notification.data.url) || './';
  e.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function(lista){
      for (var i = 0; i < lista.length; i++) {
        if ('focus' in lista[i]) return lista[i].focus();   // ya está abierta
      }
      if (self.clients.openWindow) return self.clients.openWindow(destino);
    })
  );
});
