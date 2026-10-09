/*
 * Puerta de contraseña de los tableros protegidos.
 *
 * Los datos de un tablero protegido se publican CIFRADOS (data.enc.json, AES-256 + HMAC,
 * clave derivada de la contraseña con PBKDF2-SHA256). El archivo sin cifrar no existe en el
 * sitio, así que la contraseña protege los datos de verdad, no solo la pantalla.
 * La contraseña se ingresa una vez por sesión del navegador y se descifra acá, en tu equipo.
 *
 * Uso en cada tablero:  cargarDatos('./data.json').then(data => ...)
 * Si no hay data.enc.json (tablero público) lee data.json tal cual.
 */
(function () {
  var CLAVE_SESION = 'gescom-clave-tableros';

  function aBytes(b64) {
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function unir(a, b) { var r = new Uint8Array(a.length + b.length); r.set(a, 0); r.set(b, a.length); return r; }

  function leerGuardada() { try { return sessionStorage.getItem(CLAVE_SESION) || ''; } catch (e) { return ''; } }
  function guardar(c) { try { sessionStorage.setItem(CLAVE_SESION, c); } catch (e) { /* sin almacenamiento: se vuelve a pedir */ } }
  function olvidar() { try { sessionStorage.removeItem(CLAVE_SESION); } catch (e) { /* nada */ } }

  async function descifrar(sobre, clave) {
    if (!window.crypto || !crypto.subtle) throw new Error('Este navegador no permite descifrar (se necesita https).');
    if (sobre.v !== 1) throw new Error('Formato de datos no reconocido.');
    var iter = Number(sobre.iter);
    if (!(iter >= 100000 && iter <= 5000000)) throw new Error('Parámetros de cifrado inválidos.');
    var sal = aBytes(sobre.salt), iv = aBytes(sobre.iv), ct = aBytes(sobre.ct), mac = aBytes(sobre.mac);
    var base = await crypto.subtle.importKey('raw', new TextEncoder().encode(clave), 'PBKDF2', false, ['deriveBits']);
    var bits = new Uint8Array(await crypto.subtle.deriveBits({ name: 'PBKDF2', salt: sal, iterations: iter, hash: 'SHA-256' }, base, 512));
    var kMac = await crypto.subtle.importKey('raw', bits.slice(32, 64), { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']);
    var ok = await crypto.subtle.verify('HMAC', kMac, mac, unir(iv, ct));
    if (!ok) { var e = new Error('clave'); e.claveIncorrecta = true; throw e; }
    var kEnc = await crypto.subtle.importKey('raw', bits.slice(0, 32), 'AES-CBC', false, ['decrypt']);
    var plano = await crypto.subtle.decrypt({ name: 'AES-CBC', iv: iv }, kEnc, ct);
    return JSON.parse(new TextDecoder().decode(plano));
  }

  function estilos() {
    if (document.getElementById('gate-estilos')) return;
    var s = document.createElement('style');
    s.id = 'gate-estilos';
    s.textContent =
      '#gate-velo{position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;padding:20px;background:#f5f7fa;font-family:Inter,system-ui,-apple-system,"Segoe UI",sans-serif;color:#1f2933}' +
      '#gate-caja{width:100%;max-width:360px;background:#fff;border:1px solid #e4e8ee;border-radius:12px;padding:28px 26px 24px;box-shadow:0 4px 16px rgba(16,24,40,.08)}' +
      '#gate-caja h1{margin:0 0 6px;font-size:18px;font-weight:700}' +
      '#gate-caja p{margin:0 0 18px;font-size:13.5px;line-height:1.5;color:#4a5866}' +
      '#gate-caja label{display:block;font-size:12.5px;font-weight:500;color:#7b8794;margin-bottom:6px}' +
      '#gate-clave{width:100%;box-sizing:border-box;padding:10px 12px;background:#fff;border:1px solid #d5dbe3;border-radius:8px;color:#1f2933;font-size:15px}' +
      '#gate-clave:focus{outline:2px solid #cfdcf7;border-color:#2f6fde}' +
      '#gate-ok{margin-top:14px;width:100%;padding:11px;border:0;border-radius:8px;background:#2f6fde;color:#fff;font-weight:600;font-size:14px;cursor:pointer}' +
      '#gate-ok:disabled{opacity:.6;cursor:default}' +
      '#gate-ok:focus-visible{outline:2px solid #1f2933;outline-offset:2px}' +
      '#gate-msg{min-height:18px;margin-top:12px;font-size:12.5px;color:#c2412d}';
    document.head.appendChild(s);
  }

  function pedirClave(sobre, mensajeInicial) {
    estilos();
    return new Promise(function (resolver) {
      var velo = document.createElement('div');
      velo.id = 'gate-velo';
      velo.innerHTML =
        '<form id="gate-caja" autocomplete="off">' +
        '<h1>Tablero protegido</h1>' +
        '<p>Ingresá la contraseña para ver los datos.</p>' +
        '<label for="gate-clave">Contraseña</label>' +
        '<input id="gate-clave" type="password" autocomplete="current-password" required>' +
        '<button id="gate-ok" type="submit">Ingresar</button>' +
        '<div id="gate-msg" role="alert" aria-live="polite"></div>' +
        '</form>';
      document.body.appendChild(velo);
      var form = velo.querySelector('#gate-caja'), campo = velo.querySelector('#gate-clave'),
          boton = velo.querySelector('#gate-ok'), msg = velo.querySelector('#gate-msg');
      if (mensajeInicial) msg.textContent = mensajeInicial;
      campo.focus();
      form.addEventListener('submit', async function (ev) {
        ev.preventDefault();
        var clave = campo.value;
        if (!clave) return;
        boton.disabled = true; boton.textContent = 'Descifrando…'; msg.textContent = '';
        try {
          var datos = await descifrar(sobre, clave);
          guardar(clave);
          velo.remove();
          resolver(datos);
        } catch (e) {
          boton.disabled = false; boton.textContent = 'Ingresar';
          msg.textContent = e.claveIncorrecta ? 'Contraseña incorrecta.' : ('No se pudo abrir: ' + e.message);
          campo.select();
        }
      });
    });
  }

  window.cargarDatos = async function (url) {
    var base = String(url).split('?')[0];
    var r = await fetch(base.replace(/data\.json$/, 'data.enc.json') + '?t=' + Date.now(), { cache: 'no-store' });
    if (!r.ok) {
      // sin versión cifrada: tablero público
      var plano = await fetch(base + '?t=' + Date.now(), { cache: 'no-store' });
      if (!plano.ok) throw new Error('HTTP ' + plano.status);
      return plano.json();
    }
    var sobre = await r.json();
    var guardada = leerGuardada();
    if (guardada) {
      try { return await descifrar(sobre, guardada); }
      catch (e) { if (e.claveIncorrecta) olvidar(); else throw e; }
    }
    return pedirClave(sobre, guardada ? 'La contraseña cambió: ingresá la nueva.' : '');
  };
})();
