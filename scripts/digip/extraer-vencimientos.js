// Vencidos y proximos a vencer de Digip, para valorizarlos con pull-vencidos-valor.ps1.
// Correr con javascript_exec en https://app.digipwms.com (con la sesion de Gisela iniciada)
// y guardar el resultado en un archivo FUERA del repo (ej. %TEMP%\digip-vencimientos.json):
// trae stock del deposito y el repo es publico.
// v: [codigo, descripcion, vencimiento dd/mm/aaaa, dias vencido, unidades]
// p: [codigo, descripcion, vencimiento dd/mm/aaaa, dias que quedan para despachar, unidades]
(async () => {
  const bajar = u => fetch(u, { credentials: 'same-origin' }).then(r => {
    if (r.url.includes('Login')) throw new Error('sin sesion en Digip');
    return r.text();
  });
  const filas = t => t.split('\r\n').filter(l => l.trim()).slice(1)
    .map(l => l.split(';').map(x => x.replace(/"/g, '').trim()));
  const [rv, rp] = await Promise.all([
    bajar('/Reportes/Stock/DescargarArticulosVencidos?propiedadDetalleId=0&areaId=0&articuloId=0'),
    bajar('/Reportes/Stock/DescargarArticulosProximosVencer?propiedadDetalleId=0&areaId=0&articuloId=0')
  ]);
  return JSON.stringify({
    generado: new Date().toISOString(),
    v: filas(rv).map(r => [r[0], r[1], r[4], r[5], r[7]]),
    p: filas(rp).map(r => [r[0], r[1], r[4], r[6], r[7]])
  });
})();
