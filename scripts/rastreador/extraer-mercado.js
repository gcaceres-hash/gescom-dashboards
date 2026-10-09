// Refresca scripts/rastreador/mercado.json desde el Rastreador de Lucas.
// Correr en la consola (o con javascript_exec) en
//   https://ventas.panelempresas.workers.dev/panel/rastreador   (con la sesion de Gisela iniciada)
// y guardar el resultado como scripts/rastreador/mercado.json. Despues correr pull-rastreador.ps1.
// Solo toma Lago Puelo y Elebes con precio de folleto (Pehuenia no es nuestra).
JSON.stringify({
  generado: DATA.meta.generado,
  origen: "Rastreador de precios mayoristas de Lucas (ventas.panelempresas.workers.dev/panel/rastreador), filas de Lago Puelo y Elebes con precio de folleto",
  fuentes: DATA.meta.fuentes,
  formato: "filas: [codigo Lucas (LPE-<codigo Gescom>), empresa, unidades por bulto que usa para llevar a precio por unidad, [[mayorista, precio por unidad con IVA, producto en el folleto, condiciones, link]]]",
  filas: DATA.rows.filter(r => r.e !== 'Pehuenia' && r.m && r.m.length)
    .map(r => [r.c, r.e, r.k, r.m.map(x => [x.f, x.p, x.t, x.n, x.u])])
});
