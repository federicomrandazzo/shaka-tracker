/**
 * Envía las notificaciones del tracker de Shaka.
 *
 * La planilla de Google no puede firmar mensajes push, así que le pide a esta
 * función que los mande. Corre en Vercel, junto a la app.
 *
 * Necesita tres variables de entorno (Vercel ▸ Settings ▸ Environment Variables):
 *   VAPID_PUBLIC   clave pública  (la misma que usa la app)
 *   VAPID_PRIVATE  clave privada  (NO se comparte)
 *   PUSH_SECRET    frase inventada, la misma que queda en el Apps Script
 */

const webpush = require('web-push');

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Solo POST' });
  }

  const { VAPID_PUBLIC, VAPID_PRIVATE, PUSH_SECRET } = process.env;
  if (!VAPID_PUBLIC || !VAPID_PRIVATE || !PUSH_SECRET) {
    return res.status(500).json({ ok: false, error: 'Faltan las claves en Vercel' });
  }

  let datos = req.body;
  if (typeof datos === 'string') {
    try { datos = JSON.parse(datos); } catch (e) { datos = {}; }
  }
  datos = datos || {};

  if (datos.secreto !== PUSH_SECRET) {
    return res.status(401).json({ ok: false, error: 'No autorizado' });
  }

  const subs = Array.isArray(datos.subs) ? datos.subs : [];
  if (!subs.length) return res.status(200).json({ ok: true, enviadas: 0, vencidas: [] });

  webpush.setVapidDetails('mailto:federicomrandazzo@gmail.com', VAPID_PUBLIC, VAPID_PRIVATE);

  const cuerpo = JSON.stringify({
    title: datos.title || 'Shaka · Tareas',
    body:  datos.body  || '',
    url:   datos.url   || 'https://shaka-tracker.vercel.app/',
    tag:   datos.tag   || 'shaka-tarea'
  });

  let enviadas = 0;
  const vencidas = [];   // suscripciones muertas, para que la planilla las borre

  await Promise.all(subs.map(async (s) => {
    if (!s || !s.endpoint) return;
    try {
      await webpush.sendNotification(
        { endpoint: s.endpoint, keys: { p256dh: s.p256dh, auth: s.auth } },
        cuerpo,
        { TTL: 86400 }
      );
      enviadas++;
    } catch (err) {
      // 404 o 410 = el teléfono se dio de baja o desinstaló la app
      if (err && (err.statusCode === 404 || err.statusCode === 410)) {
        vencidas.push(s.endpoint);
      }
    }
  }));

  return res.status(200).json({ ok: true, enviadas, vencidas });
};
