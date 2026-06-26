// SOP/WMS conectado a Supabase
// Proyecto: https://zeoybvogcgevzkqalsez.supabase.co

const SUPABASE_URL = "https://zeoybvogcgevzkqalsez.supabase.co";
const SUPABASE_KEY = "sb_publishable_kTeGf9zhjkyVeCblXhEtVA_oHL4L7da";
const API_BASE = SUPABASE_URL + "/rest/v1";

let datosExcel = [];
let pedidoSeleccionado = [];
let tarimasValidadas = [];
let cierreParcial = null;
let lectorCamara = null;
let escanerActivo = false;
let modoEscaneoCamara = "auto";
let pedidoActualId = null;

function $(id) {
  return document.getElementById(id);
}

function headers(extra = {}) {
  return {
    "apikey": SUPABASE_KEY,
    "Authorization": "Bearer " + SUPABASE_KEY,
    "Content-Type": "application/json",
    "Prefer": "return=representation",
    ...extra
  };
}

async function fetchConTiempoLimite(url, opciones = {}, segundos = 12) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), segundos * 1000);

  try {
    return await fetch(url, { ...opciones, signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

async function supabaseGet(path) {
  const res = await fetchConTiempoLimite(API_BASE + path, { headers: headers() });
  if (!res.ok) throw new Error(await res.text());
  return await res.json();
}

async function supabasePost(table, body) {
  const res = await fetchConTiempoLimite(`${API_BASE}/${table}`, {
    method: "POST",
    headers: headers(),
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(await res.text());
  return await res.json();
}

async function supabasePatch(table, query, body) {
  const res = await fetchConTiempoLimite(`${API_BASE}/${table}?${query}`, {
    method: "PATCH",
    headers: headers(),
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(await res.text());
  return await res.json();
}

async function supabaseDelete(table, query) {
  const res = await fetchConTiempoLimite(`${API_BASE}/${table}?${query}`, {
    method: "DELETE",
    headers: headers({ "Prefer": "return=minimal" })
  });
  if (!res.ok) throw new Error(await res.text());
}

function mostrarEstadoNube(texto) {
  const el = $("estadoNube");
  if (el) el.innerHTML = texto;
}

async function probarConexionSupabase() {
  try {
    await supabaseGet("/pedidos?select=id&limit=1");
    mostrarEstadoNube("✅ Conectado a Supabase | Datos compartidos entre PC, Zebra y celular");
  } catch (error) {
    console.error(error);
    mostrarEstadoNube("❌ Error conectando a Supabase. Revisa permisos/API Key.");
  }
}

function mostrarSeccion(id) {
  document.querySelectorAll(".seccion").forEach(sec => {
    sec.style.display = "none";
  });
  const seccion = $(id);
  if (seccion) seccion.style.display = "block";
}

async function leerExcel() {
  const archivo = $("archivoExcel")?.files[0];
  if (!archivo) {
    alert("Seleccione un archivo Excel");
    return;
  }

  const reader = new FileReader();

  reader.onload = async function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(hoja, { header: 1 });

      await procesarExcel(filas);
    } catch (error) {
      console.error(error);
      alert("No se pudo leer el Excel. Revise el archivo.");
    }
  };

  reader.onerror = function() {
    alert("El navegador no pudo abrir el archivo. En celular pruebe Chrome y permita archivos.");
  };

  reader.readAsArrayBuffer(archivo);
}

async function procesarExcel(filas) {
  datosExcel = [];
  pedidoSeleccionado = [];
  tarimasValidadas = [];
  cierreParcial = null;
  pedidoActualId = null;

  let ultimoPedido = "";
  let ultimoCliente = "";

  filas.forEach(fila => {
    const pedido = fila[2];
    const cliente = fila[3];
    const producto = fila[4];
    const cantidad = fila[5];

    if (pedido && String(pedido).trim().startsWith("S")) {
      ultimoPedido = String(pedido).trim();
    }

    if (cliente && !String(cliente).includes("Orden de venta")) {
      ultimoCliente = String(cliente).trim();
    }

    if (producto && cantidad && ultimoPedido) {
      const productoTexto = String(producto);

      if (!productoTexto.includes("Movimientos de stock")) {
        const match = productoTexto.match(/\[(.*?)\]/);

        if (match) {
          const sku = normalizarSKU(match[1]);
          const descripcion = productoTexto.replace(match[0], "").trim();
          const cantidadNumerica = convertirCantidad(cantidad);

          if (sku && cantidadNumerica > 0) {
            datosExcel.push({
              pedido: ultimoPedido,
              cliente: ultimoCliente,
              sku,
              descripcion,
              cantidadPedida: cantidadNumerica,
              cantidadValidada: 0
            });
          }
        }
      }
    }
  });

  datosExcel = consolidarSKUs(datosExcel);

  if (datosExcel.length === 0) {
    alert("No se detectaron pedidos/SKU en el Excel. Revisa el formato del archivo.");
    return;
  }

  await subirPedidosASupabase(datosExcel);
  await cargarPedidosDesdeNube();

  alert("Excel cargado a Supabase. Ya debe verse desde PC, celular y Zebra.");
}

async function subirPedidosASupabase(lineas) {
  const pedidosUnicos = [...new Set(lineas.map(x => x.pedido))];

  for (const pedido of pedidosUnicos) {
    const lineasPedido = lineas.filter(x => x.pedido === pedido);
    const cliente = lineasPedido[0]?.cliente || "";

    const existentes = await supabaseGet(`/pedidos?pedido=eq.${encodeURIComponent(pedido)}&select=id,pedido`);
    let pedidoId;

    if (existentes.length > 0) {
      const confirmar = confirm(`El pedido ${pedido} ya existe en la nube.\n\n¿Quieres reemplazar su detalle?\n\nOJO: esto borra validaciones anteriores de ese pedido.`);
      if (!confirmar) continue;

      pedidoId = existentes[0].id;

      await supabaseDelete("evidencias", `validacion_id=in.(select id from validaciones where pedido_id=${pedidoId})`).catch(() => {});
      await supabaseDelete("validaciones", `pedido_id=eq.${pedidoId}`);
      await supabaseDelete("pedido_detalle", `pedido_id=eq.${pedidoId}`);
      await supabaseDelete("cierres_parciales", `pedido_id=eq.${pedidoId}`);
      await supabasePatch("pedidos", `id=eq.${pedidoId}`, {
        cliente,
        chofer: null,
        validador: null,
        estatus: "PENDIENTE",
        fecha_cierre: null
      });
    } else {
      const creado = await supabasePost("pedidos", {
        pedido,
        cliente,
        estatus: "PENDIENTE"
      });
      pedidoId = creado[0].id;
    }

    const detalle = lineasPedido.map(x => ({
      pedido_id: pedidoId,
      sku: normalizarSKU(x.sku),
      descripcion: x.descripcion || "",
      cantidad_pedida: Number(x.cantidadPedida) || 0,
      cantidad_validada: 0
    }));

    if (detalle.length > 0) {
      await supabasePost("pedido_detalle", detalle);
    }
  }
}

async function cargarPedidosDesdeNube() {
  try {
    const pedidos = await supabaseGet("/pedidos?select=*&order=fecha_creacion.desc");
    const detalles = await supabaseGet("/pedido_detalle?select=*");

    datosExcel = detalles.map(d => {
      const p = pedidos.find(x => x.id === d.pedido_id) || {};
      return {
        id: d.id,
        pedidoId: d.pedido_id,
        pedido: p.pedido || "",
        cliente: p.cliente || "",
        sku: d.sku,
        descripcion: d.descripcion || "",
        cantidadPedida: Number(d.cantidad_pedida || 0),
        cantidadValidada: Number(d.cantidad_validada || 0),
        estatus: p.estatus || "PENDIENTE"
      };
    });

    mostrarPedidosDetectados();
    mostrarEstadoNube(`✅ Nube actualizada | Pedidos: ${pedidos.length} | SKU: ${detalles.length}`);
    return pedidos;
  } catch (error) {
    console.error(error);
    alert("No se pudieron cargar pedidos desde Supabase.");
    return [];
  }
}

function convertirCantidad(valor) {
  if (typeof valor === "number") return valor;
  return Number(String(valor).replace(/,/g, "").replace(/\s/g, "")) || 0;
}

function normalizarSKU(valor) {
  return String(valor || "").trim().replace(/^\[/, "").replace(/\]$/, "").toUpperCase();
}

function extraerSKUDesdeScan(valor) {
  const datos = extraerDatosDesdeTexto(valor);
  if (datos.sku) return normalizarSKU(datos.sku);

  const texto = String(valor || "").trim();
  const entreCorchetes = texto.match(/\[(.*?)\]/);
  if (entreCorchetes) return normalizarSKU(entreCorchetes[1]);

  const skuEtiqueta = texto.match(/SKU[:=\s]+([A-Za-z0-9._-]+)/i);
  if (skuEtiqueta) return normalizarSKU(skuEtiqueta[1]);

  return normalizarSKU(texto);
}

function extraerDatosDesdeTexto(valor) {
  const textoOriginal = String(valor || "").trim();
  const datos = { sku: "", lote: "", caducidad: "", cantidad: "", idEtiqueta: "" };
  if (!textoOriginal) return datos;

  // =====================================================
  // NORMALIZADOR UNIVERSAL PARA QR / SCANNER ZEBRA / CÁMARA
  // Soporta datos pegados:
  // SKU=MAAG18LOTE=A0626/1532CAD=02-jun-2027ID=ENV-20260623-000042
  // Y datos separados por líneas, |, ;, coma, JSON o GS1.
  // =====================================================

  const textoSinSaltos = textoOriginal
    .replace(/\r/g, "")
    .replace(/\n/g, "")
    .trim();

  // 1) Formato pegado con claves. Este es el caso del láser Zebra.
  const mapaPegado = extraerCamposPegadosPorClave(textoSinSaltos);
  if (mapaPegado.SKU || mapaPegado.LOTE || mapaPegado.CAD || mapaPegado.CADUCIDAD || mapaPegado.CANTIDAD || mapaPegado.CANT || mapaPegado.QTY || mapaPegado.ID || mapaPegado.FOLIO) {
    datos.sku = normalizarSKU(mapaPegado.SKU || mapaPegado.CODIGO || mapaPegado.CLAVE || "");
    datos.lote = String(mapaPegado.LOTE || mapaPegado.LOT || mapaPegado.BATCH || "").trim();
    datos.caducidad = normalizarFecha(mapaPegado.CAD || mapaPegado.CADUCIDAD || mapaPegado.EXP || mapaPegado.VENCE || "");
    datos.cantidad = String(mapaPegado.CANTIDAD || mapaPegado.CANT || mapaPegado.QTY || mapaPegado.PIEZAS || mapaPegado.PZAS || "")
      .replace(/[^0-9.]/g, "");
    datos.idEtiqueta = normalizarIdEtiqueta(mapaPegado.ID || mapaPegado.FOLIO || "");

    // Si al menos encontró algo útil, regresar de inmediato.
    if (datos.sku || datos.lote || datos.caducidad || datos.cantidad || datos.idEtiqueta) return datos;
  }

  // 2) JSON.
  try {
    const json = JSON.parse(textoOriginal);
    datos.sku = normalizarSKU(json.sku || json.SKU || json.codigo || json.codigoProducto || json.producto || "");
    datos.lote = String(json.lote || json.LOTE || json.batch || json.Batch || "").trim();
    datos.caducidad = normalizarFecha(json.caducidad || json.CADUCIDAD || json.exp || json.EXP || json.fechaCaducidad || "");
    datos.cantidad = String(json.cantidad || json.CANTIDAD || json.qty || json.QTY || json.piezas || "").replace(/[^0-9.]/g, "");
    datos.idEtiqueta = normalizarIdEtiqueta(json.id || json.ID || json.folio || json.FOLIO || json.idEtiqueta || json.ID_ETIQUETA || "");
    return datos;
  } catch (_) {}

  // 3) Formato con separadores normales.
  const texto = textoOriginal
    .replace(/\r/g, "\n")
    .replace(/\|/g, "\n")
    .replace(/;/g, "\n")
    .replace(/,/g, "\n");

  datos.sku = extraerValorPorClaves(texto, ["SKU", "CODIGO", "CÓDIGO", "PRODUCTO", "ITEM", "ARTICULO", "ARTÍCULO", "CLAVE"]);
  datos.lote = extraerValorPorClaves(texto, ["LOTE", "LOT", "BATCH"]);
  datos.caducidad = normalizarFecha(extraerValorPorClaves(texto, ["CADUCIDAD", "CAD", "EXP", "VENCE", "VENCIMIENTO", "FECHA CADUCIDAD"]));
  datos.cantidad = extraerValorPorClaves(texto, ["CANTIDAD", "CANT", "QTY", "PIEZAS", "PZAS", "PZA", "PCS"]);
  datos.idEtiqueta = normalizarIdEtiqueta(extraerValorPorClaves(texto, ["ID", "FOLIO", "ID ETIQUETA", "ID_ETIQUETA"]));

  // 4) GS1 con identificadores.
  if (!datos.sku) {
    const ai240 = textoOriginal.match(/\(240\)([^()]+)/);
    const ai241 = textoOriginal.match(/\(241\)([^()]+)/);
    const ai91 = textoOriginal.match(/\(91\)([^()]+)/);
    datos.sku = ai240?.[1] || ai241?.[1] || ai91?.[1] || "";
  }

  if (!datos.lote) {
    const ai10 = textoOriginal.match(/\(10\)([^()]+)/);
    datos.lote = ai10?.[1] || "";
  }

  if (!datos.caducidad) {
    const ai17 = textoOriginal.match(/\(17\)(\d{6})/);
    if (ai17) datos.caducidad = convertirFechaGS1(ai17[1]);
  }

  if (!datos.cantidad) {
    const ai30 = textoOriginal.match(/\(30\)(\d+)/);
    const ai37 = textoOriginal.match(/\(37\)(\d+)/);
    datos.cantidad = ai30?.[1] || ai37?.[1] || "";
  }

  // 5) SKU entre corchetes o SKU plano.
  if (!datos.sku) {
    const entreCorchetes = textoOriginal.match(/\[(.*?)\]/);
    if (entreCorchetes) datos.sku = entreCorchetes[1];
  }

  if (!datos.sku && /^[A-Za-z0-9._-]{3,40}$/.test(textoOriginal)) datos.sku = textoOriginal;

  datos.sku = normalizarSKU(datos.sku);
  datos.lote = String(datos.lote || "").trim();
  datos.caducidad = normalizarFecha(datos.caducidad);
  datos.cantidad = datos.cantidad ? String(datos.cantidad).replace(/[^0-9.]/g, "") : "";
  datos.idEtiqueta = normalizarIdEtiqueta(datos.idEtiqueta);

  return datos;
}

function extraerCamposPegadosPorClave(texto) {
  const resultado = {};
  const limpio = String(texto || "").trim();
  if (!limpio) return resultado;

  // Claves que pueden venir pegadas sin separador.
  const claves = [
    "SKU", "CODIGO", "CÓDIGO", "CLAVE",
    "LOTE", "LOT", "BATCH",
    "CADUCIDAD", "CAD", "EXP", "VENCE", "VENCIMIENTO",
    "CANTIDAD", "CANT", "QTY", "PIEZAS", "PZAS",
    "ID", "FOLIO"
  ];

  const patronClaves = claves
    .map(c => c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .join("|");

  const re = new RegExp(`(${patronClaves})\\s*=`, "gi");
  const encontrados = [];
  let m;

  while ((m = re.exec(limpio)) !== null) {
    encontrados.push({
      clave: m[1].toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""),
      inicioValor: re.lastIndex,
      inicioClave: m.index
    });
  }

  for (let i = 0; i < encontrados.length; i++) {
    const actual = encontrados[i];
    const siguiente = encontrados[i + 1];
    const fin = siguiente ? siguiente.inicioClave : limpio.length;
    const valor = limpio.slice(actual.inicioValor, fin).trim();
    resultado[actual.clave] = valor;
  }

  return resultado;
}

function normalizarIdEtiqueta(valor) {
  return String(valor || "")
    .trim()
    .replace(/^(ID|FOLIO|ID ETIQUETA|ID_ETIQUETA)\s*[:=\-#]?\s*/i, "")
    .replace(/\s+/g, "")
    .toUpperCase();
}

function extraerValorPorClaves(texto, claves) {
  for (const clave of claves) {
    const claveSegura = clave.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const patron = new RegExp(`(?:^|\\n|\\s)${claveSegura}\\s*[:=\\-#]\\s*([^\\n]+)`, "i");
    const match = texto.match(patron);
    if (match && match[1]) return match[1].trim();
  }
  return "";
}

function normalizarFecha(valor) {
  let texto = String(valor || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  if (!texto) return "";

  // Limpia prefijos comunes de QR/etiqueta.
  texto = texto
    .replace(/^(CADUCIDAD|CAD|EXP|VENCE|VENCIMIENTO|FECHA CADUCIDAD)\s*[:=\-#]?\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();

  // yyyy-mm-dd o yyyy/mm/dd
  let match = texto.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (match) {
    const yyyy = match[1];
    const mm = match[2].padStart(2, "0");
    const dd = match[3].padStart(2, "0");
    return fechaValidaISO(yyyy, mm, dd);
  }

  // dd/mm/yyyy, dd-mm-yyyy o dd.mm.yyyy
  match = texto.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (match) {
    const dd = match[1].padStart(2, "0");
    const mm = match[2].padStart(2, "0");
    const yyyy = match[3];
    return fechaValidaISO(yyyy, mm, dd);
  }

  // dd MES yyyy: 19 MAYO 2026, 19-MAY-2026, 19MAYO2026
  const meses = {
    ENERO: "01", ENE: "01",
    FEBRERO: "02", FEB: "02",
    MARZO: "03", MAR: "03",
    ABRIL: "04", ABR: "04",
    MAYO: "05", MAY: "05",
    JUNIO: "06", JUN: "06",
    JULIO: "07", JUL: "07",
    AGOSTO: "08", AGO: "08",
    SEPTIEMBRE: "09", SETIEMBRE: "09", SEP: "09",
    OCTUBRE: "10", OCT: "10",
    NOVIEMBRE: "11", NOV: "11",
    DICIEMBRE: "12", DIC: "12"
  };

  match = texto.match(/^(\d{1,2})\s*[\/\-. ]?\s*([A-ZÑ]+)\s*[\/\-. ]?\s*(\d{4})$/);
  if (match) {
    const dd = match[1].padStart(2, "0");
    const mm = meses[match[2].replace(/Ñ/g, "N")];
    const yyyy = match[3];
    if (mm) return fechaValidaISO(yyyy, mm, dd);
  }

  // GS1: yymmdd
  match = texto.match(/^(\d{2})(\d{2})(\d{2})$/);
  if (match) return convertirFechaGS1(texto);

  // Busca una fecha dentro de un texto más largo.
  match = texto.match(/(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{4})/);
  if (match) return normalizarFecha(match[1]);

  match = texto.match(/(\d{1,2}\s*[\/\-. ]?\s*[A-ZÑ]+\s*[\/\-. ]?\s*\d{4})/);
  if (match) return normalizarFecha(match[1]);

  return "";
}

function fechaValidaISO(yyyy, mm, dd) {
  const iso = `${yyyy}-${mm}-${dd}`;
  const fecha = new Date(`${iso}T00:00:00`);
  if (
    Number.isNaN(fecha.getTime()) ||
    fecha.getUTCFullYear() !== Number(yyyy) ||
    fecha.getUTCMonth() + 1 !== Number(mm) ||
    fecha.getUTCDate() !== Number(dd)
  ) {
    return "";
  }
  return iso;
}

function convertirFechaGS1(valor) {
  const texto = String(valor || "").trim();
  if (!/^\d{6}$/.test(texto)) return "";
  const yy = Number(texto.slice(0, 2));
  const mm = texto.slice(2, 4);
  const dd = texto.slice(4, 6);
  const yyyy = yy >= 70 ? `19${String(yy).padStart(2, "0")}` : `20${String(yy).padStart(2, "0")}`;
  return `${yyyy}-${mm}-${dd}`;
}

function aplicarDatosEscaneados(valor, modo = "auto") {
  const texto = String(valor || "").trim();
  if (!texto) return;
  const datos = extraerDatosDesdeTexto(texto);

  if ($("datosLeidosQR")) $("datosLeidosQR").value = texto;

  if (modo === "sku") {
    if (datos.sku) $("skuEscaneado").value = datos.sku;
    enfocarCampoDespuesDeSKU();
    return;
  }

  if (modo === "lote") {
    if (datos.lote || texto) $("loteEscaneado").value = datos.lote || texto;
    $("caducidadEscaneada").focus();
    return;
  }

  if (datos.sku) $("skuEscaneado").value = datos.sku;
  if (datos.lote) $("loteEscaneado").value = datos.lote;
  if (datos.caducidad) $("caducidadEscaneada").value = normalizarFecha(datos.caducidad);
  if (datos.cantidad) $("cantidadTarima").value = datos.cantidad;
  if (datos.idEtiqueta && $("idEtiquetaEscaneada")) $("idEtiquetaEscaneada").value = datos.idEtiqueta;

  if (!datos.sku && texto) $("skuEscaneado").value = extraerSKUDesdeScan(texto);

  if (!$("loteEscaneado").value) $("loteEscaneado").focus();
  else if (!$("caducidadEscaneada").value) $("caducidadEscaneada").focus();
  else if (!$("cantidadTarima").value) {
    $("cantidadTarima").focus();
    $("cantidadTarima").select?.();
  }
  else if ($("fotoTarima1Camara")) $("fotoTarima1Camara").focus();
}

function enfocarCampoDespuesDeSKU() {
  if (!$("loteEscaneado").value) $("loteEscaneado").focus();
  else if (!$("caducidadEscaneada").value) $("caducidadEscaneada").focus();
  else if (!$("cantidadTarima").value) $("cantidadTarima").focus();
  else if ($("fotoTarima1Camara")) $("fotoTarima1Camara").focus();
}

function consolidarSKUs(datos) {
  const mapa = {};
  datos.forEach(item => {
    const clave = `${item.pedido}|${item.sku}`;
    if (!mapa[clave]) {
      mapa[clave] = {
        pedido: item.pedido,
        cliente: item.cliente,
        sku: item.sku,
        descripcion: item.descripcion,
        cantidadPedida: 0,
        cantidadValidada: 0
      };
    }
    mapa[clave].cantidadPedida += Number(item.cantidadPedida) || 0;
  });
  return Object.values(mapa);
}

function mostrarPedidosDetectados() {
  const pedidosUnicos = [...new Set(datosExcel
    .filter(x => ["", "PENDIENTE", "EN PROCESO"].includes(String(x.estatus || "PENDIENTE").toUpperCase()))
    .map(x => x.pedido))];

  $("resumenExcel").innerHTML = `
    <h4>Pedidos disponibles</h4>
    <p>SKU consolidados: <b>${datosExcel.length}</b></p>
    <p>Pedidos encontrados: <b>${pedidosUnicos.length}</b></p>
  `;

  let opciones = `<option value="">Seleccione un pedido</option>`;
  pedidosUnicos.forEach(pedido => {
    opciones += `<option value="${escaparHTML(pedido)}">${escaparHTML(pedido)}</option>`;
  });

  $("selectorPedido").innerHTML = `
    <h4>Seleccione el pedido a validar</h4>
    <select id="pedidoDetectado">${opciones}</select>
    <button onclick="seleccionarPedido()">Cargar pedido seleccionado</button>
    <button onclick="verDashboardPedidos()">Ver dashboard de pedidos</button>
    <button onclick="verHistorial()">Ver historial</button>
  `;
}

async function seleccionarPedido() {
  const pedido = $("pedidoDetectado")?.value;
  if (!pedido) {
    alert("Seleccione un pedido");
    return;
  }
  await cargarPedidoDesdeDashboard(pedido);
}

async function cargarPedidoDesdeDashboard(pedido) {
  try {
    const pedidos = await supabaseGet(`/pedidos?pedido=eq.${encodeURIComponent(pedido)}&select=*`);
    if (pedidos.length === 0) {
      alert("No se encontró el pedido en Supabase");
      return;
    }

    const p = pedidos[0];
    pedidoActualId = p.id;
    cierreParcial = null;

    const detalles = await supabaseGet(`/pedido_detalle?pedido_id=eq.${p.id}&select=*&order=id.asc`);
    const cierres = await supabaseGet(`/cierres_parciales?pedido_id=eq.${p.id}&select=*&order=fecha_cierre.desc&limit=1`);

    if (cierres.length > 0) {
      cierreParcial = {
        pedido: p.pedido,
        cliente: p.cliente || "",
        motivo: cierres[0].motivo || "",
        autorizo: cierres[0].autorizo || "",
        comentario: cierres[0].comentario || "",
        fechaHora: new Date(cierres[0].fecha_cierre).toLocaleString()
      };
    }

    pedidoSeleccionado = detalles.map(d => ({
      id: d.id,
      pedidoId: p.id,
      pedido: p.pedido,
      cliente: p.cliente || "",
      sku: d.sku,
      descripcion: d.descripcion || "",
      cantidadPedida: Number(d.cantidad_pedida || 0),
      cantidadValidada: Number(d.cantidad_validada || 0)
    }));

    await cargarTarimasPedido(p.id, p.pedido);

    $("pedido").value = p.pedido;
    $("cliente").value = p.cliente || "";
    $("chofer").value = p.chofer || "";
    $("validador").value = p.validador || "";

    actualizarAvance();
    mostrarSeccion("pasoDatos");
  } catch (error) {
    console.error(error);
    alert("No se pudo abrir el pedido desde Supabase.");
  }
}

async function cargarTarimasPedido(pedidoId, pedidoTexto) {
  const validaciones = await supabaseGet(`/validaciones?pedido_id=eq.${pedidoId}&select=*&order=fecha_validacion.asc`);
  const evidencias = await supabaseGet(`/evidencias?select=*`);

  tarimasValidadas = validaciones.map(v => {
    const linea = pedidoSeleccionado.find(x => normalizarSKU(x.sku) === normalizarSKU(v.sku)) || {};
    const ev = evidencias.filter(e => e.validacion_id === v.id).sort((a,b) => a.id - b.id);
    return {
      validacionId: v.id,
      idEtiqueta: v.id_etiqueta || "",
      pedido: pedidoTexto,
      cliente: linea.cliente || $("cliente")?.value || "",
      sku: v.sku,
      descripcion: linea.descripcion || "",
      lote: v.lote,
      caducidad: v.caducidad,
      cantidad: Number(v.cantidad || 0),
      fechaHora: new Date(v.fecha_validacion).toLocaleString(),
      foto1Nombre: ev[0]?.nombre_archivo || "",
      foto2Nombre: ev[1]?.nombre_archivo || "",
      foto1Base64: ev[0]?.url_archivo || "",
      foto2Base64: ev[1]?.url_archivo || ""
    };
  });
}

async function validarDatosGenerales() {
  const pedido = $("pedido").value.trim();
  const cliente = $("cliente").value.trim();
  const chofer = $("chofer").value.trim();
  const validador = $("validador").value.trim();

  if (!pedido || !cliente || !chofer || !validador) {
    alert("Debe completar Pedido, Cliente, Chofer y Validador");
    return;
  }

  if (!pedidoActualId) {
    const pedidos = await supabaseGet(`/pedidos?pedido=eq.${encodeURIComponent(pedido)}&select=id`);
    pedidoActualId = pedidos[0]?.id || null;
  }

  if (pedidoActualId) {
    await supabasePatch("pedidos", `id=eq.${pedidoActualId}`, {
      cliente,
      chofer,
      validador,
      estatus: obtenerEstatusPedido(pedido)
    });
  }

  actualizarAvance();
  mostrarSeccion("pasoValidacion");
  enfocarSKU();
}

function obtenerRegistroEtiquetasLocal() {
  try {
    return JSON.parse(localStorage.getItem("sopwms_ids_etiquetas_validadas") || "[]");
  } catch (_) {
    return [];
  }
}

function guardarRegistroEtiquetaLocal(registro) {
  if (!registro?.idEtiqueta) return;
  const registros = obtenerRegistroEtiquetasLocal();
  if (!registros.some(r => normalizarIdEtiqueta(r.idEtiqueta) === normalizarIdEtiqueta(registro.idEtiqueta))) {
    registros.push(registro);
    localStorage.setItem("sopwms_ids_etiquetas_validadas", JSON.stringify(registros.slice(-10000)));
  }
}

async function buscarEtiquetaDuplicada(idEtiqueta) {
  const id = normalizarIdEtiqueta(idEtiqueta);
  if (!id) return null;

  const local = obtenerRegistroEtiquetasLocal().find(r => normalizarIdEtiqueta(r.idEtiqueta) === id);
  if (local) return { origen: "local", ...local };

  // Si Supabase ya tiene una columna id_etiqueta, se usa para bloqueo global entre equipos.
  // Si la columna aún no existe, no se rompe la app: queda activo el bloqueo local.
  try {
    const encontrados = await supabaseGet(`/validaciones?id_etiqueta=eq.${encodeURIComponent(id)}&select=*,pedidos(pedido,cliente)&limit=1`);
    if (encontrados && encontrados.length > 0) {
      const v = encontrados[0];
      return {
        origen: "nube",
        idEtiqueta: id,
        pedido: v.pedidos?.pedido || v.pedido || "",
        cliente: v.pedidos?.cliente || "",
        sku: v.sku || "",
        fechaHora: v.fecha_validacion ? new Date(v.fecha_validacion).toLocaleString() : ""
      };
    }
  } catch (error) {
    console.warn("Validación global de ID no disponible; se usará bloqueo local.", error);
  }

  return null;
}

async function guardarTarima() {
  // =====================================================
  // VALIDACIÓN FINAL CORREGIDA
  // El QR solo llena SKU, LOTE y CADUCIDAD.
  // La CANTIDAD debe capturarla manualmente el operador.
  // Antes el mensaje decía que faltaba todo, aunque realmente
  // solo faltara la cantidad. Ahora indica exactamente qué falta.
  // =====================================================
  const campoSKU = $("skuEscaneado");
  const campoLote = $("loteEscaneado");
  const campoCaducidad = $("caducidadEscaneada");
  const campoCantidad = $("cantidadTarima");

  const sku = extraerSKUDesdeScan(campoSKU?.value || "");
  const lote = String(campoLote?.value || "").trim();
  const caducidad = String(campoCaducidad?.value || "").trim();
  const cantidadTexto = String(campoCantidad?.value || "").trim().replace(",", ".");
  const cantidad = Number(cantidadTexto);
  let idEtiqueta = normalizarIdEtiqueta($("idEtiquetaEscaneada")?.value || "");
  if (!idEtiqueta) {
    idEtiqueta = normalizarIdEtiqueta(extraerDatosDesdeTexto($("datosLeidosQR")?.value || campoSKU?.value || "").idEtiqueta);
    if (idEtiqueta && $("idEtiquetaEscaneada")) $("idEtiquetaEscaneada").value = idEtiqueta;
  }

  const faltantes = [];
  if (!sku) faltantes.push("SKU");
  if (!lote) faltantes.push("lote");
  if (!caducidad) faltantes.push("caducidad");
  if (!cantidadTexto) faltantes.push("cantidad por tarima");

  if (faltantes.length > 0) {
    alert("Falta capturar: " + faltantes.join(", "));

    if (!sku && campoSKU) campoSKU.focus();
    else if (!lote && campoLote) campoLote.focus();
    else if (!caducidad && campoCaducidad) campoCaducidad.focus();
    else if (!cantidadTexto && campoCantidad) campoCantidad.focus();

    return;
  }

  if (Number.isNaN(cantidad) || cantidad <= 0) {
    alert("La cantidad por tarima debe ser un número mayor a cero");
    if (campoCantidad) campoCantidad.focus();
    return;
  }

  const foto1 = obtenerArchivoEvidencia("fotoTarima1Camara", "fotoTarima1Galeria");
  const foto2 = obtenerArchivoEvidencia("fotoTarima2Camara", "fotoTarima2Galeria");

  if (!foto1 && !foto2) {
    alert("Debe agregar al menos una evidencia fotográfica");
    return;
  }

  const lineaPedido = pedidoSeleccionado.find(x => normalizarSKU(x.sku) === sku);

  if (!lineaPedido) {
    alert("ERROR: El SKU no pertenece a este pedido");
    enfocarSKU();
    return;
  }

  if (cierreParcial && cierreParcial.pedido === lineaPedido.pedido) {
    alert("Este pedido ya fue cerrado como parcial. No se pueden agregar más tarimas.");
    return;
  }

  if (idEtiqueta) {
    const duplicada = await buscarEtiquetaDuplicada(idEtiqueta);
    if (duplicada) {
      alert(
        "ERROR: Esta etiqueta ya fue escaneada.\n\n" +
        "ID etiqueta: " + idEtiqueta + "\n" +
        "Pedido: " + (duplicada.pedido || "No disponible") + "\n" +
        "SKU: " + (duplicada.sku || "No disponible") + "\n" +
        "Fecha: " + (duplicada.fechaHora || "No disponible") + "\n\n" +
        "No se permite validar la misma etiqueta en este u otro pedido."
      );
      limpiarCapturaTarima();
      return;
    }
  }

  const nuevoTotal = Number(lineaPedido.cantidadValidada || 0) + cantidad;

  if (nuevoTotal > Number(lineaPedido.cantidadPedida || 0)) {
    alert(
      "ERROR: Cantidad excedida\n\n" +
      "SKU: " + sku + "\n" +
      "Pedido: " + lineaPedido.cantidadPedida + "\n" +
      "Validado actual: " + lineaPedido.cantidadValidada + "\n" +
      "Intentas agregar: " + cantidad + "\n" +
      "Total resultante: " + nuevoTotal
    );
    return;
  }

  let foto1Base64 = "";
  let foto2Base64 = "";

  try {
    foto1Base64 = await convertirImagenABase64(foto1);
    foto2Base64 = await convertirImagenABase64(foto2);
  } catch (error) {
    console.error(error);
    alert("No se pudo procesar la imagen. Intente con otra foto.");
    return;
  }

  try {
    const usuario = $("validador").value.trim() || "Sin usuario";

    let validacionCreada;
    const payloadValidacion = {
      pedido_id: pedidoActualId || lineaPedido.pedidoId,
      sku: lineaPedido.sku,
      lote,
      caducidad,
      cantidad,
      usuario
    };

    try {
      validacionCreada = await supabasePost("validaciones", idEtiqueta ? { ...payloadValidacion, id_etiqueta: idEtiqueta } : payloadValidacion);
    } catch (error) {
      console.warn("No se pudo guardar id_etiqueta en Supabase; se guarda validación normal y bloqueo local.", error);
      validacionCreada = await supabasePost("validaciones", payloadValidacion);
    }

    const validacionId = validacionCreada[0].id;

    const evidencias = [];
    if (foto1Base64) evidencias.push({
      validacion_id: validacionId,
      nombre_archivo: foto1 ? foto1.name : "foto1.jpg",
      url_archivo: foto1Base64
    });
    if (foto2Base64) evidencias.push({
      validacion_id: validacionId,
      nombre_archivo: foto2 ? foto2.name : "foto2.jpg",
      url_archivo: foto2Base64
    });

    if (evidencias.length > 0) await supabasePost("evidencias", evidencias);

    await supabasePatch("pedido_detalle", `id=eq.${lineaPedido.id}`, {
      cantidad_validada: nuevoTotal
    });

    lineaPedido.cantidadValidada = nuevoTotal;

    const lineaExcel = datosExcel.find(x =>
      x.pedido === lineaPedido.pedido &&
      normalizarSKU(x.sku) === normalizarSKU(lineaPedido.sku)
    );
    if (lineaExcel) lineaExcel.cantidadValidada = nuevoTotal;

    tarimasValidadas.push({
      validacionId,
      idEtiqueta,
      pedido: lineaPedido.pedido,
      cliente: lineaPedido.cliente,
      sku: lineaPedido.sku,
      descripcion: lineaPedido.descripcion,
      lote,
      caducidad,
      cantidad,
      fechaHora: new Date().toLocaleString(),
      foto1Nombre: foto1 ? foto1.name : "",
      foto2Nombre: foto2 ? foto2.name : "",
      foto1Base64,
      foto2Base64
    });

    if (idEtiqueta) guardarRegistroEtiquetaLocal({
      idEtiqueta,
      pedido: lineaPedido.pedido,
      cliente: lineaPedido.cliente,
      sku: lineaPedido.sku,
      lote,
      caducidad,
      fechaHora: new Date().toLocaleString()
    });

    const estatus = pedidoCompletado() ? "COMPLETADO" : "EN PROCESO";
    await supabasePatch("pedidos", `id=eq.${pedidoActualId || lineaPedido.pedidoId}`, {
      estatus,
      fecha_cierre: estatus === "COMPLETADO" ? new Date().toISOString() : null
    });

    limpiarFormularioTarima();
    actualizarAvance();

    if (pedidoCompletado()) {
      alert("PEDIDO COMPLETADO CORRECTAMENTE");
      mostrarSeccion("pasoAvance");
    } else {
      alert("Tarima guardada correctamente en Supabase");
    }
  } catch (error) {
    console.error(error);
    alert("No se pudo guardar en Supabase. Revisa conexión o permisos.");
  }
}

function limpiarFormularioTarima() {
  $("skuEscaneado").value = "";
  $("loteEscaneado").value = "";
  $("caducidadEscaneada").value = "";
  $("cantidadTarima").value = "";
  limpiarEvidencias();
  if ($("datosLeidosQR")) $("datosLeidosQR").value = "";
  if ($("idEtiquetaEscaneada")) $("idEtiquetaEscaneada").value = "";
  enfocarSKU();
}

function limpiarCapturaTarima() {
  $("skuEscaneado").value = "";
  $("loteEscaneado").value = "";
  $("caducidadEscaneada").value = "";
  $("cantidadTarima").value = "";
  if ($("datosLeidosQR")) $("datosLeidosQR").value = "";
  if ($("idEtiquetaEscaneada")) $("idEtiquetaEscaneada").value = "";
  enfocarSKU();
}

function enfocarSKU() {
  setTimeout(() => {
    const campo = $("skuEscaneado");
    if (campo) campo.focus();
  }, 100);
}

function actualizarAvance() {
  let totalPedido = 0;
  let totalValidado = 0;
  let html = "";

  pedidoSeleccionado.forEach(item => {
    totalPedido += Number(item.cantidadPedida || 0);
    totalValidado += Number(item.cantidadValidada || 0);
    const pendiente = Number(item.cantidadPedida || 0) - Number(item.cantidadValidada || 0);

    html += `
      <div class="linea-avance">
        <b>${escaparHTML(item.sku)}</b><br>
        ${escaparHTML(item.descripcion)}<br>
        Pedido: ${item.cantidadPedida} |
        Validado: ${item.cantidadValidada} |
        Pendiente: ${pendiente}
      </div>
    `;
  });

  const porcentaje = totalPedido > 0 ? Math.round((totalValidado / totalPedido) * 100) : 0;

  const datosParcial = cierreParcial && cierreParcial.pedido === ($("pedido")?.value || "")
    ? `
      <div class="linea-avance">
        <b>CIERRE PARCIAL</b><br>
        Motivo: ${escaparHTML(cierreParcial.motivo)}<br>
        Autorizó: ${escaparHTML(cierreParcial.autorizo)}<br>
        Comentario: ${escaparHTML(cierreParcial.comentario || "Sin comentario")}<br>
        Fecha: ${escaparHTML(cierreParcial.fechaHora)}
      </div>
    `
    : "";

  $("avancePedido").innerHTML = `
    <h4>Avance general: ${porcentaje}%</h4>
    <p>Total pedido: <b>${totalPedido}</b></p>
    <p>Total validado: <b>${totalValidado}</b></p>
    <p>Total pendiente: <b>${totalPedido - totalValidado}</b></p>
    <p>Estatus: <b>${obtenerEstatusPedido($("pedido")?.value || "")}</b></p>
    ${datosParcial}
    <hr>
    ${html}
  `;
}

function pedidoCompletado() {
  return pedidoSeleccionado.length > 0 &&
    pedidoSeleccionado.every(item => Number(item.cantidadValidada) === Number(item.cantidadPedida));
}

function obtenerEstatusPedido(pedido) {
  if (cierreParcial && cierreParcial.pedido === pedido) return "COMPLETADO PARCIAL";

  const lineas = datosExcel.filter(x => x.pedido === pedido);
  if (lineas.length === 0 && pedidoSeleccionado.length > 0) {
    return pedidoCompletado() ? "COMPLETADO" : "EN PROCESO";
  }
  if (lineas.length === 0) return "PENDIENTE";

  const totalPedido = lineas.reduce((suma, item) => suma + Number(item.cantidadPedida || 0), 0);
  const totalValidado = lineas.reduce((suma, item) => suma + Number(item.cantidadValidada || 0), 0);

  if (totalPedido > 0 && totalValidado === totalPedido) return "COMPLETADO";
  if (totalValidado > 0) return "EN PROCESO";
  return "PENDIENTE";
}

async function cerrarPedidoParcial() {
  if (pedidoSeleccionado.length === 0) {
    alert("Primero seleccione un pedido");
    return;
  }

  const pedido = $("pedido").value.trim();
  if (pedidoCompletado()) {
    alert("El pedido ya está completo. No requiere cierre parcial.");
    return;
  }

  const totalValidado = pedidoSeleccionado.reduce((suma, item) => suma + Number(item.cantidadValidada || 0), 0);
  if (totalValidado <= 0) {
    const confirmarSinCarga = confirm("No hay cantidades validadas. ¿Aun así desea cerrar parcial por faltante?");
    if (!confirmarSinCarga) return;
  }

  const motivo = prompt("Motivo del cierre parcial por faltante:");
  if (!motivo || !motivo.trim()) {
    alert("El motivo es obligatorio");
    return;
  }

  const autorizo = prompt("Nombre de quien autoriza el cierre parcial:");
  if (!autorizo || !autorizo.trim()) {
    alert("El autorizador es obligatorio");
    return;
  }

  const comentario = prompt("Comentario adicional:") || "";

  try {
    await supabasePost("cierres_parciales", {
      pedido_id: pedidoActualId || pedidoSeleccionado[0].pedidoId,
      motivo: motivo.trim(),
      autorizo: autorizo.trim(),
      comentario: comentario.trim()
    });

    await supabasePatch("pedidos", `id=eq.${pedidoActualId || pedidoSeleccionado[0].pedidoId}`, {
      estatus: "COMPLETADO PARCIAL",
      fecha_cierre: new Date().toISOString()
    });

    cierreParcial = {
      pedido,
      cliente: $("cliente").value.trim(),
      motivo: motivo.trim(),
      autorizo: autorizo.trim(),
      comentario: comentario.trim(),
      fechaHora: new Date().toLocaleString()
    };

    actualizarAvance();
    alert("Pedido cerrado como COMPLETADO PARCIAL en Supabase");
    mostrarSeccion("pasoAvance");
  } catch (error) {
    console.error(error);
    alert("No se pudo cerrar parcial en Supabase.");
  }
}

function verResumenValidacion() {
  if (pedidoSeleccionado.length === 0) {
    alert("Primero seleccione un pedido");
    return;
  }
  actualizarAvance();
  $("resumenValidacion").innerHTML = $("avancePedido").innerHTML;
  mostrarSeccion("pasoResumen");
}

async function verDashboardPedidos() {
  try {
    await cargarPedidosDesdeNube();

    if (datosExcel.length === 0) {
      $("dashboardPedidos").innerHTML = "<p>No hay pedidos en Supabase.</p>";
      mostrarDashboardKPIsGlobal(0, 0, 0, 0, 0, 0, 0);
      mostrarSeccion("pasoDashboard");
      return;
    }

    const pedidosUnicos = [...new Set(datosExcel.map(x => x.pedido))];

    let completados = 0;
    let parciales = 0;
    let enProceso = 0;
    let pendientes = 0;
    let totalPedidoGlobal = 0;
    let totalValidadoGlobal = 0;
    let html = "";

    pedidosUnicos.forEach(pedido => {
      const lineas = datosExcel.filter(x => x.pedido === pedido);
      const cliente = lineas[0]?.cliente || "";
      const totalPedido = lineas.reduce((suma, item) => suma + Number(item.cantidadPedida || 0), 0);
      const totalValidado = lineas.reduce((suma, item) => suma + Number(item.cantidadValidada || 0), 0);
      const pendiente = totalPedido - totalValidado;
      const avance = totalPedido > 0 ? Math.round((totalValidado / totalPedido) * 100) : 0;
      const estatus = lineas[0]?.estatus || obtenerEstatusPedido(pedido);

      totalPedidoGlobal += totalPedido;
      totalValidadoGlobal += totalValidado;

      if (estatus === "COMPLETADO") completados++;
      else if (estatus === "COMPLETADO PARCIAL") parciales++;
      else if (estatus === "EN PROCESO") enProceso++;
      else pendientes++;

      html += crearDetallePedidoDashboardHTML({
        pedido,
        cliente,
        totalPedido,
        totalValidado,
        pendiente,
        avance,
        estatus
      });
    });

    const avanceGlobal = totalPedidoGlobal > 0
      ? Math.round((totalValidadoGlobal / totalPedidoGlobal) * 100)
      : 0;

    const porcentajeCompletados = pedidosUnicos.length > 0
      ? Math.round((completados / pedidosUnicos.length) * 100)
      : 0;

    const porcentajeEnProceso = pedidosUnicos.length > 0
      ? Math.round((enProceso / pedidosUnicos.length) * 100)
      : 0;

    const porcentajePendientes = pedidosUnicos.length > 0
      ? Math.round((pendientes / pedidosUnicos.length) * 100)
      : 0;

    window.dashboardPedidosDetalleHTML = html;
    window.dashboardPedidosStats = {
      avance: { porcentaje: avanceGlobal, totalPedidos: pedidosUnicos.length, totalPedidoGlobal, totalValidadoGlobal },
      completados: { porcentaje: porcentajeCompletados, total: completados },
      enProceso: { porcentaje: porcentajeEnProceso, total: enProceso },
      pendientes: { porcentaje: porcentajePendientes, total: pendientes }
    };

    mostrarDashboardKPIsGlobal(
      avanceGlobal,
      porcentajeCompletados,
      porcentajeEnProceso,
      porcentajePendientes,
      pedidosUnicos.length,
      totalPedidoGlobal,
      totalValidadoGlobal
    );

    $("dashboardPedidos").innerHTML = `<div id="infoDashboardGrafico" class="dashboard-info" style="display:none;"></div>`;
    mostrarSeccion("pasoDashboard");
  } catch (error) {
    console.error(error);
    alert("No se pudo abrir dashboard desde Supabase.");
  }
}


function alternarDetalleDashboard() {
  const detalle = $("detalleDashboardPedidos");
  if (!detalle) return;
  detalle.style.display = detalle.style.display === "none" ? "block" : "none";
}

function mostrarDashboardKPIsGlobal(avance, completados, enProceso, pendientes, totalPedidos, totalPedidoGlobal, totalValidadoGlobal) {
  const contenedor = $("dashboardKPIs");
  if (!contenedor) return;

  contenedor.innerHTML = `
    <div class="dashboard-gauges">
      ${crearGaugeHTML("Avance general", avance, `${totalValidadoGlobal} / ${totalPedidoGlobal}`, "avance")}
      ${crearGaugeHTML("Pedidos completos", completados, `${completados}%`, "completados")}
      ${crearGaugeHTML("En proceso", enProceso, `${enProceso}%`, "enProceso")}
      ${crearGaugeHTML("Pendientes", pendientes, `${pendientes}%`, "pendientes")}
    </div>
  `;
}

function crearGaugeHTML(titulo, porcentaje, subtitulo, tipo) {
  const valor = Math.max(0, Math.min(100, Number(porcentaje) || 0));
  return `
    <button type="button" class="gauge-card gauge-click" onclick="mostrarInfoDashboardGrafico('${escaparAtributo(tipo)}')">
      <div class="gauge" style="--valor:${valor}">
        <div class="gauge-centro">
          <span>${valor}%</span>
        </div>
      </div>
      <div class="gauge-titulo">${escaparHTML(titulo)}</div>
      <div class="gauge-subtitulo">${escaparHTML(subtitulo || "")}</div>
    </button>
  `;
}

function crearDetallePedidoDashboardHTML(item) {
  return `
    <div class="linea-avance">
      <b>Pedido:</b> ${escaparHTML(item.pedido)}<br>
      <b>Cliente:</b> ${escaparHTML(item.cliente)}<br>
      <b>Total pedido:</b> ${item.totalPedido}<br>
      <b>Total validado:</b> ${item.totalValidado}<br>
      <b>Pendiente:</b> ${item.pendiente}<br>
      <b>Avance:</b> ${item.avance}%<br>
      <b>Estatus:</b> ${escaparHTML(item.estatus)}<br>
      <button onclick="cargarPedidoDesdeDashboard('${escaparAtributo(item.pedido)}')">Abrir pedido</button>
      <button onclick="generarPDFDesdePedido('${escaparAtributo(item.pedido)}')">PDF</button>
    </div>
  `;
}

function mostrarInfoDashboardGrafico(tipo) {
  const panel = $("infoDashboardGrafico");
  if (!panel) return;

  const stats = window.dashboardPedidosStats || {};
  let titulo = "Detalle";
  let resumen = "";

  if (tipo === "avance") {
    const s = stats.avance || {};
    titulo = "Detalle de avance general";
    resumen = `Pedido total: <b>${s.totalPedidoGlobal || 0}</b><br>Validado: <b>${s.totalValidadoGlobal || 0}</b><br>Pedidos: <b>${s.totalPedidos || 0}</b><br>Avance: <b>${s.porcentaje || 0}%</b>`;
  } else if (tipo === "completados") {
    const s = stats.completados || {};
    titulo = "Detalle de pedidos completos";
    resumen = `Pedidos completos: <b>${s.total || 0}</b><br>Porcentaje: <b>${s.porcentaje || 0}%</b>`;
  } else if (tipo === "enProceso") {
    const s = stats.enProceso || {};
    titulo = "Detalle de pedidos en proceso";
    resumen = `Pedidos en proceso: <b>${s.total || 0}</b><br>Porcentaje: <b>${s.porcentaje || 0}%</b>`;
  } else if (tipo === "pendientes") {
    const s = stats.pendientes || {};
    titulo = "Detalle de pedidos pendientes";
    resumen = `Pedidos pendientes: <b>${s.total || 0}</b><br>Porcentaje: <b>${s.porcentaje || 0}%</b>`;
  }

  panel.style.display = "block";
  panel.innerHTML = `
    <h4>${titulo}</h4>
    <p>${resumen}</p>
    <div>${window.dashboardPedidosDetalleHTML || ""}</div>
  `;
}

function guardarEstado() {
  // Ya no usamos localStorage para operación. Supabase es la fuente de verdad.
}

function recuperarEstado() {
  // Ya no recuperamos proceso local. Al abrir se cargan pedidos desde Supabase.
}

function limpiarMemoriaPedido() {
  pedidoSeleccionado = [];
  tarimasValidadas = [];
  cierreParcial = null;
  pedidoActualId = null;
  alert("Pantalla local limpiada. Los datos de Supabase no se borraron.");
  location.reload();
}

function borrarTodoLocal() {
  pedidoSeleccionado = [];
  tarimasValidadas = [];
  cierreParcial = null;
  pedidoActualId = null;
  $("resumenExcel").innerHTML = "";
  $("selectorPedido").innerHTML = "";
  alert("Pantalla local limpiada. Los datos en nube siguen intactos.");
}

function obtenerSemana(fecha) {
  const f = new Date(fecha);
  const primerDia = new Date(f.getFullYear(), 0, 1);
  const dias = Math.floor((f - primerDia) / (24 * 60 * 60 * 1000));
  const semana = Math.ceil((dias + primerDia.getDay() + 1) / 7);
  return `${f.getFullYear()}-W${semana}`;
}

function guardarHistorialLigero() {
  // Historial ahora es Supabase: pedidos + validaciones + cierres.
}

async function verHistorial(buscar = false) {
  try {
    if (!buscar) {
      if ($("buscarHistorial")) $("buscarHistorial").value = "";
      if ($("filtroPeriodo")) $("filtroPeriodo").value = "todos";
      if ($("historialValidaciones")) $("historialValidaciones").innerHTML = "";
      mostrarSeccion("pasoHistorial");
      return;
    }

    const pedidos = await supabaseGet("/pedidos?select=*&order=fecha_creacion.desc");
    const validaciones = await supabaseGet("/validaciones?select=*");

    const busqueda = $("buscarHistorial")?.value.toLowerCase() || "";
    const periodo = $("filtroPeriodo")?.value || "todos";
    const hoy = new Date();
    const diaActual = hoy.toISOString().slice(0, 10);
    const mesActual = hoy.toISOString().slice(0, 7);
    const semanaActual = obtenerSemana(hoy);

    let filtrado = pedidos.filter(p => {
      const coincideBusqueda =
        String(p.pedido || "").toLowerCase().includes(busqueda) ||
        String(p.cliente || "").toLowerCase().includes(busqueda);

      const fecha = p.fecha_creacion ? new Date(p.fecha_creacion) : new Date();
      const dia = fecha.toISOString().slice(0, 10);
      const mes = fecha.toISOString().slice(0, 7);
      const semana = obtenerSemana(fecha);

      let coincidePeriodo = true;
      if (periodo === "dia") coincidePeriodo = dia === diaActual;
      if (periodo === "semana") coincidePeriodo = semana === semanaActual;
      if (periodo === "mes") coincidePeriodo = mes === mesActual;

      return coincideBusqueda && coincidePeriodo;
    });

    if (filtrado.length === 0) {
      $("historialValidaciones").innerHTML = `<p>No hay registros encontrados.</p>`;
      mostrarSeccion("pasoHistorial");
      return;
    }

    let html = "";
    filtrado.forEach(p => {
      const vals = validaciones.filter(v => v.pedido_id === p.id);
      html += `
        <div class="linea-avance">
          <b>Pedido:</b> ${escaparHTML(p.pedido)}<br>
          <b>Cliente:</b> ${escaparHTML(p.cliente || "")}<br>
          <b>Chofer:</b> ${escaparHTML(p.chofer || "")}<br>
          <b>Validador:</b> ${escaparHTML(p.validador || "")}<br>
          <b>Fecha:</b> ${p.fecha_creacion ? new Date(p.fecha_creacion).toLocaleString() : ""}<br>
          <b>Estatus:</b> ${escaparHTML(p.estatus || "")}<br>
          <b>Tarimas validadas:</b> ${vals.length}<br><br>
          <button onclick="generarPDFDesdePedido('${escaparAtributo(p.pedido)}')">Generar PDF</button>
        </div>
      `;
    });

    $("historialValidaciones").innerHTML = html;
    mostrarSeccion("pasoHistorial");
  } catch (error) {
    console.error(error);
    alert("No se pudo consultar historial en Supabase.");
  }
}

function generarPDF() {
  if (pedidoSeleccionado.length === 0) {
    alert("Primero seleccione un pedido");
    return;
  }

  generarPDFPedido(
    pedidoSeleccionado,
    tarimasValidadas,
    {
      pedido: $("pedido").value,
      cliente: $("cliente").value,
      chofer: $("chofer").value,
      validador: $("validador").value,
      cierreParcial: cierreParcial && cierreParcial.pedido === $("pedido").value ? cierreParcial : null
    }
  );
}

function generarPDFPedido(lineasPedido, tarimas, datos) {
  if (!window.jspdf || !window.jspdf.jsPDF) {
    alert("No se cargó la librería PDF. Revise conexión a internet.");
    return;
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  const fecha = new Date().toLocaleString();
  const totalPedido = lineasPedido.reduce((s, x) => s + Number(x.cantidadPedida || 0), 0);
  const totalValidado = lineasPedido.reduce((s, x) => s + Number(x.cantidadValidada || 0), 0);
  const pendiente = totalPedido - totalValidado;

  const estatusPDF = datos.cierreParcial ? "COMPLETADO PARCIAL" : pendiente === 0 ? "COMPLETADO" : "INCOMPLETO";

  let y = 15;
  doc.setFontSize(16);
  doc.text("VALIDACIÓN DE CARGA VS PEDIDO", 14, y);

  y += 10;
  doc.setFontSize(10);
  y = escribirLineaPDF(doc, `Pedido: ${datos.pedido}`, 14, y);
  y = escribirLineaPDF(doc, `Cliente: ${datos.cliente}`, 14, y);
  y = escribirLineaPDF(doc, `Chofer: ${datos.chofer}`, 14, y);
  y = escribirLineaPDF(doc, `Validador: ${datos.validador}`, 14, y);
  y = escribirLineaPDF(doc, `Fecha generación PDF: ${fecha}`, 14, y);

  y += 6;
  doc.setFontSize(12);
  doc.text("RESUMEN DEL PEDIDO", 14, y);
  y += 8;
  doc.setFontSize(10);
  y = escribirLineaPDF(doc, `Total pedido original: ${totalPedido}`, 14, y);
  y = escribirLineaPDF(doc, `Total validado: ${totalValidado}`, 14, y);
  y = escribirLineaPDF(doc, `Total pendiente: ${pendiente}`, 14, y);
  y = escribirLineaPDF(doc, `Estatus: ${estatusPDF}`, 14, y);

  if (datos.cierreParcial) {
    y += 4;
    doc.setFontSize(12);
    doc.text("CIERRE PARCIAL POR FALTANTE", 14, y);
    y += 8;
    doc.setFontSize(10);
    y = escribirLineaPDF(doc, `Motivo: ${datos.cierreParcial.motivo}`, 14, y);
    y = escribirLineaPDF(doc, `Autorizó: ${datos.cierreParcial.autorizo}`, 14, y);
    y = escribirLineaPDF(doc, `Comentario: ${datos.cierreParcial.comentario || "Sin comentario"}`, 14, y);
    y = escribirLineaPDF(doc, `Fecha cierre parcial: ${datos.cierreParcial.fechaHora}`, 14, y);
  }

  y += 6;
  doc.setFontSize(12);
  doc.text("DETALLE DE SKU", 14, y);
  y += 8;
  doc.setFontSize(9);

  lineasPedido.forEach(item => {
    if (y > 260) {
      doc.addPage();
      y = 15;
    }
    const pendienteSku = Number(item.cantidadPedida || 0) - Number(item.cantidadValidada || 0);
    y = escribirLineaPDF(doc, `SKU: ${item.sku}`, 14, y, 180, 5);
    y = escribirLineaPDF(doc, `Descripción: ${item.descripcion}`, 14, y, 180, 5);
    y = escribirLineaPDF(doc, `Pedido: ${item.cantidadPedida} | Validado: ${item.cantidadValidada} | Pendiente: ${pendienteSku}`, 14, y, 180, 5);
    y += 3;
  });

  const tarimasPedido = tarimas.filter(t => t.pedido === datos.pedido);

  tarimasPedido.forEach((t, index) => {
    doc.addPage();
    y = 15;
    doc.setFontSize(12);
    doc.text(`REGISTRO VALIDADO ${index + 1}`, 14, y);
    y += 8;
    doc.setFontSize(9);
    y = escribirLineaPDF(doc, `SKU: ${t.sku}`, 14, y, 180, 5);
    y = escribirLineaPDF(doc, `Descripción: ${t.descripcion}`, 14, y, 180, 5);
    y = escribirLineaPDF(doc, `Lote: ${t.lote}`, 14, y, 180, 5);
    y = escribirLineaPDF(doc, `Caducidad: ${t.caducidad}`, 14, y, 180, 5);
    y = escribirLineaPDF(doc, `Cantidad: ${t.cantidad}`, 14, y, 180, 5);
    y = escribirLineaPDF(doc, `Fecha/hora: ${t.fechaHora}`, 14, y, 180, 5);

    y += 4;
    doc.text("Evidencia 1:", 14, y);
    y += 5;
    y = insertarImagenPDF(doc, t.foto1Base64, 14, y);

    y += 5;
    doc.text("Evidencia 2:", 14, y);
    y += 5;
    y = insertarImagenPDF(doc, t.foto2Base64, 14, y);
  });

  doc.addPage();
  y = 20;
  doc.setFontSize(12);
  doc.text("CIERRE DE VALIDACIÓN", 14, y);
  y += 10;
  doc.setFontSize(10);
  y = escribirLineaPDF(doc, `Resultado final: ${estatusPDF}`, 14, y);
  y += 20;
  doc.text("Firma / Nombre del validador:", 14, y);
  y += 15;
  doc.line(14, y, 100, y);
  y += 6;
  doc.text(datos.validador || "Sin validador", 14, y);

  const clienteLimpio = limpiarTextoArchivo(datos.cliente);
  const pedidoLimpio = limpiarTextoArchivo(datos.pedido);
  doc.save(`${clienteLimpio}_${pedidoLimpio}.pdf`);
  alert("PDF generado correctamente");
}

function escribirLineaPDF(doc, texto, x, y, ancho = 180, salto = 6) {
  const lineas = doc.splitTextToSize(String(texto || ""), ancho);
  doc.text(lineas, x, y);
  return y + (lineas.length * salto);
}

function insertarImagenPDF(doc, imagenBase64, x, y) {
  if (!imagenBase64) {
    doc.text("Sin imagen", x, y);
    return y + 8;
  }

  try {
    const tipo = obtenerTipoImagen(imagenBase64);
    doc.addImage(imagenBase64, tipo, x, y, 85, 65);
    return y + 72;
  } catch (error) {
    console.error(error);
    doc.text("No se pudo insertar la imagen", x, y);
    return y + 8;
  }
}

function obtenerTipoImagen(base64) {
  if (base64.startsWith("data:image/png")) return "PNG";
  if (base64.startsWith("data:image/webp")) return "WEBP";
  return "JPEG";
}

function limpiarTextoArchivo(texto) {
  return String(texto || "").replace(/[^a-zA-Z0-9_-]/g, "_").substring(0, 40);
}

async function generarPDFDesdePedido(pedido) {
  await cargarPedidoDesdeDashboard(pedido);
  generarPDF();
}

async function generarPDFDesdeHistorial(pedido) {
  await generarPDFDesdePedido(pedido);
}

function convertirImagenABase64(archivo) {
  return new Promise((resolve, reject) => {
    if (!archivo) {
      resolve("");
      return;
    }

    const reader = new FileReader();

    reader.onload = function(evento) {
      const img = new Image();

      img.onload = function() {
        const maxAncho = 900;
        const escala = Math.min(1, maxAncho / img.width);
        const ancho = Math.round(img.width * escala);
        const alto = Math.round(img.height * escala);
        const canvas = document.createElement("canvas");
        canvas.width = ancho;
        canvas.height = alto;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, ancho, alto);
        resolve(canvas.toDataURL("image/jpeg", 0.60));
      };

      img.onerror = reject;
      img.src = evento.target.result;
    };

    reader.onerror = reject;
    reader.readAsDataURL(archivo);
  });
}

function abrirCamaraEvidencia(inputId) {
  const input = $(inputId);
  if (!input) {
    alert("No se encontró el campo de cámara.");
    return;
  }
  input.click();
}

function abrirGaleriaEvidencia(inputId) {
  const input = $(inputId);
  if (!input) {
    alert("No se encontró el campo de galería.");
    return;
  }
  input.click();
}

function obtenerArchivoEvidencia(idCamara, idGaleria) {
  const archivoCamara = $(idCamara)?.files?.[0] || null;
  const archivoGaleria = $(idGaleria)?.files?.[0] || null;
  return archivoCamara || archivoGaleria || null;
}

function mostrarNombreEvidencia(inputId, contenedorId) {
  const archivo = $(inputId)?.files?.[0] || null;
  const contenedor = $(contenedorId);

  if (!contenedor) return;

  if (!archivo) {
    contenedor.textContent = "Sin evidencia seleccionada.";
    return;
  }

  contenedor.textContent = "✅ Evidencia seleccionada: " + archivo.name;
}

function limpiarInputArchivo(id) {
  const input = $(id);
  if (input) input.value = "";
}

function limpiarEvidencias() {
  limpiarInputArchivo("fotoTarima1Camara");
  limpiarInputArchivo("fotoTarima1Galeria");
  limpiarInputArchivo("fotoTarima2Camara");
  limpiarInputArchivo("fotoTarima2Galeria");

  if ($("nombreFotoTarima1")) $("nombreFotoTarima1").textContent = "Sin evidencia 1 seleccionada.";
  if ($("nombreFotoTarima2")) $("nombreFotoTarima2").textContent = "Sin evidencia 2 seleccionada.";
}

function iniciarEscanerCamara(modo = "auto") {
  if (typeof Html5Qrcode === "undefined") {
    alert("No se cargó la librería de escaneo. Revise conexión a internet.");
    return;
  }

  if (escanerActivo) return;

  modoEscaneoCamara = modo;
  lectorCamara = new Html5Qrcode("lectorCamara");

  const formatos = typeof Html5QrcodeSupportedFormats !== "undefined"
    ? [
        Html5QrcodeSupportedFormats.QR_CODE,
        Html5QrcodeSupportedFormats.CODE_128,
        Html5QrcodeSupportedFormats.CODE_39,
        Html5QrcodeSupportedFormats.EAN_13,
        Html5QrcodeSupportedFormats.EAN_8,
        Html5QrcodeSupportedFormats.UPC_A,
        Html5QrcodeSupportedFormats.UPC_E,
        Html5QrcodeSupportedFormats.ITF,
        Html5QrcodeSupportedFormats.DATA_MATRIX
      ]
    : undefined;

  const config = { fps: 10, qrbox: { width: 280, height: 180 } };
  if (formatos) config.formatsToSupport = formatos;

  lectorCamara.start(
    { facingMode: "environment" },
    config,
    codigo => procesarCodigoCamara(codigo),
    () => {}
  ).then(() => {
    escanerActivo = true;
  }).catch(error => {
    console.error(error);
    alert("No se pudo abrir la cámara. Revise permisos del navegador.");
  });
}

function procesarCodigoCamara(codigo) {
  aplicarDatosEscaneados(codigo, modoEscaneoCamara);
  detenerEscanerCamara();
}

function detenerEscanerCamara() {
  if (!lectorCamara || !escanerActivo) return;

  lectorCamara.stop()
    .then(() => {
      lectorCamara.clear();
      escanerActivo = false;
      lectorCamara = null;
    })
    .catch(error => {
      console.error(error);
      escanerActivo = false;
      lectorCamara = null;
    });
}

function escaparHTML(valor) {
  return String(valor ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escaparAtributo(valor) {
  return String(valor ?? "")
    .replace(/\\/g, "\\\\")
    .replace(/'/g, "\\'")
    .replace(/"/g, "&quot;");
}



let modoAmbiente = "PRUEBAS";

async function cargarModoAmbiente() {
  try {
    const registros = await supabaseGet("/ambiente?select=*&nombre=eq.SISTEMA&limit=1");

    if (registros && registros.length > 0) {
      modoAmbiente = String(registros[0].modo || "PRUEBAS").toUpperCase();
    } else {
      modoAmbiente = "PRUEBAS";
    }
  } catch (error) {
    console.warn("No se pudo leer tabla ambiente. Usando modo PRUEBAS por seguridad.", error);
    modoAmbiente = "PRUEBAS";
  }

  pintarPanelAmbiente();
}

function pintarPanelAmbiente() {
  const panel = $("panelAmbiente");
  const texto = $("textoAmbiente");
  const boton = $("btnLimpiarPruebas");
  const nota = $("notaAmbiente");

  if (!panel || !texto || !boton) return;

  panel.style.display = "block";

  if (modoAmbiente === "PRODUCCION") {
    panel.className = "ambiente ambiente-produccion";
    texto.textContent = "🔒 Modo: PRODUCCIÓN";
    boton.style.display = "none";
    if (nota) nota.textContent = "La limpieza masiva está deshabilitada para proteger datos reales.";
    return;
  }

  panel.className = "ambiente ambiente-pruebas";
  texto.textContent = "🧪 Modo: PRUEBAS";
  boton.style.display = "block";
  if (nota) nota.textContent = "Puedes borrar todos los pedidos, validaciones y evidencias de prueba.";
}

async function limpiarDatosPrueba() {
  if (modoAmbiente !== "PRUEBAS") {
    alert("El sistema está en PRODUCCIÓN. No se permite borrar información.");
    return;
  }

  const confirmar1 = confirm(
    "ATENCIÓN: Esto borrará TODOS los datos de prueba en Supabase:\n\n" +
    "- Pedidos\n" +
    "- Detalle de pedidos\n" +
    "- Validaciones\n" +
    "- Evidencias\n" +
    "- Cierres parciales\n\n" +
    "¿Deseas continuar?"
  );

  if (!confirmar1) return;

  const confirmar2 = prompt("Para confirmar escribe BORRAR");

  if (String(confirmar2 || "").trim().toUpperCase() !== "BORRAR") {
    alert("Limpieza cancelada.");
    return;
  }

  try {
    mostrarEstadoNube("⏳ Limpiando datos de prueba...");

    await supabaseDelete("evidencias", "id=gt.0").catch(error => console.warn("evidencias", error));
    await supabaseDelete("validaciones", "id=gt.0").catch(error => console.warn("validaciones", error));
    await supabaseDelete("cierres_parciales", "id=gt.0").catch(error => console.warn("cierres_parciales", error));
    await supabaseDelete("pedido_detalle", "id=gt.0").catch(error => console.warn("pedido_detalle", error));
    await supabaseDelete("pedidos", "id=gt.0").catch(error => console.warn("pedidos", error));

    datosExcel = [];
    pedidoSeleccionado = [];
    tarimasValidadas = [];
    cierreParcial = null;
    pedidoActualId = null;

    if ($("resumenExcel")) $("resumenExcel").innerHTML = "";
    if ($("selectorPedido")) $("selectorPedido").innerHTML = "";
    if ($("dashboardPedidos")) $("dashboardPedidos").innerHTML = "";
    if ($("historialValidaciones")) $("historialValidaciones").innerHTML = "";
    if ($("avancePedido")) $("avancePedido").innerHTML = "";

    limpiarFormularioTarima();
    mostrarSeccion("pasoExcel");

    mostrarEstadoNube("✅ Datos de prueba eliminados. Supabase quedó limpio.");
    alert("Datos de prueba eliminados correctamente.");

    await cargarPedidosDesdeNube();
    pintarPanelAmbiente();
  } catch (error) {
    console.error(error);
    mostrarEstadoNube("❌ Error limpiando datos de prueba.");
    alert("No se pudieron limpiar los datos de prueba. Revisa permisos en Supabase.");
  }
}


// =====================================================
// TEMA CLARO / OSCURO - UI V2
// Solo cambia apariencia. No modifica lógica de operación.
// =====================================================
function aplicarTemaGuardado() {
  const tema = localStorage.getItem("sopwms_tema") || "claro";
  document.body.classList.toggle("tema-oscuro", tema === "oscuro");
  const boton = $("btnTema");
  if (boton) {
    boton.textContent = tema === "oscuro" ? "☀️" : "🌙";
    boton.title = tema === "oscuro" ? "Cambiar a modo claro" : "Cambiar a modo oscuro";
  }
}

function alternarTema() {
  const temaActual = document.body.classList.contains("tema-oscuro") ? "oscuro" : "claro";
  const temaNuevo = temaActual === "oscuro" ? "claro" : "oscuro";
  localStorage.setItem("sopwms_tema", temaNuevo);
  aplicarTemaGuardado();
}

document.addEventListener("DOMContentLoaded", async () => {
  aplicarTemaGuardado();

  try {
    mostrarEstadoNube("⏳ Conectando con Supabase...");
    await probarConexionSupabase();
    await cargarModoAmbiente();
    await cargarPedidosDesdeNube();
  } catch (error) {
    console.error("Error inicial cargando Supabase:", error);
    mostrarEstadoNube("❌ No conectó con Supabase. Recarga con Ctrl+F5 o revisa internet.");
  }

  const skuInput = $("skuEscaneado");
  const loteInput = $("loteEscaneado");

  // =====================================================
  // LECTOR FÍSICO ZEBRA / LÁSER COMO TECLADO
  // NO toca el flujo de cámara. La cámara sigue entrando por
  // procesarCodigoCamara() -> aplicarDatosEscaneados().
  //
  // Problema corregido:
  // El lector físico manda el QR en varias líneas:
  // SKU=MAAG18
  // LOTE=A0626/1532
  // CAD=02-jun-2027
  // ID=ENV-20260618-000039
  //
  // Antes el Enter de la primera línea podía procesar solo SKU.
  // Ahora esperamos a que termine el lector y procesamos el bloque completo.
  // =====================================================
  let timerLecturaZebra = null;
  let ultimoQRProcesadoZebra = "";

  function pareceQRDeEtiqueta(texto) {
    const t = String(texto || "").toUpperCase();
    return (
      /SKU\s*=/.test(t) &&
      /(LOTE|LOT|BATCH)\s*=/.test(t) &&
      /(CAD|CADUCIDAD|EXP|VENCE|VENCIMIENTO)\s*=/.test(t)
    );
  }

  function programarLecturaZebra() {
    if (!skuInput) return;

    clearTimeout(timerLecturaZebra);

    timerLecturaZebra = setTimeout(() => {
      const texto = String(skuInput.value || "").trim();
      if (!texto) return;

      if (pareceQRDeEtiqueta(texto)) {
        // IMPORTANTE:
        // No bloquear QR repetidos. Si se escanea el mismo QR otra vez,
        // el lector ya dejó el texto completo dentro del campo SKU.
        // Si regresamos sin procesar, el usuario ve todo el QR pegado en SKU.
        ultimoQRProcesadoZebra = texto;
        aplicarDatosEscaneados(texto, "auto");
        return;
      }

      // Refuerzo para lectores Zebra/láser que pegan el QR en una sola línea,
      // con separadores distintos o incluso junto a un valor anterior del SKU.
      const datosDetectados = extraerDatosDesdeTexto(texto);
      if (datosDetectados.sku && (datosDetectados.lote || datosDetectados.caducidad || datosDetectados.cantidad)) {
        aplicarDatosEscaneados(texto, "auto");
        return;
      }

      // Escaneo/manual de solo SKU. No procesa si todavía parece lectura parcial.
      if (!/[=\n\r{};|:]/.test(texto)) {
        skuInput.value = extraerSKUDesdeScan(texto);
      }
    }, 450);
  }

  if (skuInput) {
    skuInput.addEventListener("input", programarLecturaZebra);

    skuInput.addEventListener("keydown", event => {
      if (event.key === "Enter") {
        // En textarea dejamos que entre el salto de línea para capturar el QR completo.
        // Solo evitamos acciones raras de formularios, aunque aquí no hay form.
        setTimeout(programarLecturaZebra, 0);
      }
    });

    skuInput.addEventListener("paste", () => {
      setTimeout(programarLecturaZebra, 0);
    });

    skuInput.addEventListener("change", programarLecturaZebra);
  }

  if (loteInput) {
    loteInput.addEventListener("keydown", event => {
      if (event.key === "Enter") {
        event.preventDefault();
        $("caducidadEscaneada").focus();
      }
    });
  }
});
// =====================================================
// EXPORTACIÓN DE BASE DE DATOS A EXCEL PARA REPORTES
// No modifica la operación: solo consulta Supabase y descarga .xlsx
// =====================================================
function formatearFechaExcel(valor) {
  if (!valor) return "";
  const fecha = new Date(valor);
  if (Number.isNaN(fecha.getTime())) return String(valor);
  return fecha.toLocaleString("es-MX");
}

function numeroExcel(valor) {
  const n = Number(valor || 0);
  return Number.isFinite(n) ? n : 0;
}

function nombreArchivoSeguro(valor) {
  return String(valor || "")
    .replace(/[^a-zA-Z0-9_-]/g, "_")
    .replace(/_+/g, "_")
    .substring(0, 80);
}

async function exportarBaseDatosExcel() {
  try {
    if (typeof XLSX === "undefined") {
      alert("No se encontró la librería de Excel. Revise conexión a internet o el CDN XLSX.");
      return;
    }

    const confirmar = confirm(
      "Se descargará un Excel con la base de datos actual para reportes:\n\n" +
      "• Resumen de pedidos\n" +
      "• Detalle por SKU\n" +
      "• Validaciones / tarimas\n" +
      "• Cierres parciales\n\n" +
      "¿Deseas continuar?"
    );
    if (!confirmar) return;

    const pedidos = await supabaseGet("/pedidos?select=*&order=fecha_creacion.desc");
    const detalles = await supabaseGet("/pedido_detalle?select=*");
    const validaciones = await supabaseGet("/validaciones?select=*&order=fecha_validacion.desc");

    let cierres = [];
    try {
      cierres = await supabaseGet("/cierres_parciales?select=*&order=fecha_cierre.desc");
    } catch (error) {
      console.warn("No se pudo consultar cierres_parciales", error);
      cierres = [];
    }

    let evidencias = [];
    try {
      evidencias = await supabaseGet("/evidencias?select=id,validacion_id,nombre_archivo");
    } catch (error) {
      console.warn("No se pudo consultar evidencias", error);
      evidencias = [];
    }

    const pedidoPorId = new Map(pedidos.map(p => [p.id, p]));
    const detallePorPedido = new Map();
    detalles.forEach(d => {
      if (!detallePorPedido.has(d.pedido_id)) detallePorPedido.set(d.pedido_id, []);
      detallePorPedido.get(d.pedido_id).push(d);
    });

    const evidenciaPorValidacion = new Map();
    evidencias.forEach(e => {
      if (!evidenciaPorValidacion.has(e.validacion_id)) evidenciaPorValidacion.set(e.validacion_id, []);
      evidenciaPorValidacion.get(e.validacion_id).push(e.nombre_archivo || `Evidencia ${e.id}`);
    });

    const validacionesPorPedido = new Map();
    validaciones.forEach(v => {
      if (!validacionesPorPedido.has(v.pedido_id)) validacionesPorPedido.set(v.pedido_id, []);
      validacionesPorPedido.get(v.pedido_id).push(v);
    });

    const resumenPedidos = pedidos.map(p => {
      const det = detallePorPedido.get(p.id) || [];
      const vals = validacionesPorPedido.get(p.id) || [];
      const totalPedido = det.reduce((s, d) => s + numeroExcel(d.cantidad_pedida), 0);
      const totalValidado = det.reduce((s, d) => s + numeroExcel(d.cantidad_validada), 0);
      const pendiente = totalPedido - totalValidado;
      const avance = totalPedido > 0 ? totalValidado / totalPedido : 0;

      return {
        Pedido: p.pedido || "",
        Cliente: p.cliente || "",
        Chofer: p.chofer || "",
        Validador: p.validador || "",
        Estatus: p.estatus || "",
        "Fecha creación": formatearFechaExcel(p.fecha_creacion),
        "Fecha cierre": formatearFechaExcel(p.fecha_cierre),
        "SKU líneas": det.length,
        "Tarimas validadas": vals.length,
        "Cantidad pedida": totalPedido,
        "Cantidad validada": totalValidado,
        Pendiente: pendiente,
        "Avance %": Math.round(avance * 10000) / 100
      };
    });

    const detallePedido = detalles.map(d => {
      const p = pedidoPorId.get(d.pedido_id) || {};
      const pedida = numeroExcel(d.cantidad_pedida);
      const validada = numeroExcel(d.cantidad_validada);
      const pendiente = pedida - validada;
      return {
        Pedido: p.pedido || "",
        Cliente: p.cliente || "",
        SKU: d.sku || "",
        Descripción: d.descripcion || "",
        "Cantidad pedida": pedida,
        "Cantidad validada": validada,
        Pendiente: pendiente,
        "Estatus línea": pendiente <= 0 ? "COMPLETA" : validada > 0 ? "PARCIAL" : "PENDIENTE"
      };
    });

    const validacionesTarimas = validaciones.map(v => {
      const p = pedidoPorId.get(v.pedido_id) || {};
      const det = (detallePorPedido.get(v.pedido_id) || []).find(d => normalizarSKU(d.sku) === normalizarSKU(v.sku)) || {};
      const evs = evidenciaPorValidacion.get(v.id) || [];
      return {
        Pedido: p.pedido || "",
        Cliente: p.cliente || "",
        SKU: v.sku || "",
        Descripción: det.descripcion || "",
        Lote: v.lote || "",
        Caducidad: v.caducidad || "",
        Cantidad: numeroExcel(v.cantidad),
        Usuario: v.usuario || "",
        "Fecha validación": formatearFechaExcel(v.fecha_validacion),
        "ID etiqueta": v.id_etiqueta || "",
        "ID validación": v.id || "",
        "Evidencias": evs.length,
        "Archivos evidencia": evs.join(" | ")
      };
    });

    const cierresParciales = cierres.map(c => {
      const p = pedidoPorId.get(c.pedido_id) || {};
      return {
        Pedido: p.pedido || c.pedido || "",
        Cliente: p.cliente || "",
        Motivo: c.motivo || c.observaciones || "",
        Usuario: c.usuario || "",
        "Fecha cierre": formatearFechaExcel(c.fecha_cierre || c.created_at)
      };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumenPedidos), "Resumen pedidos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detallePedido), "Detalle SKU");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(validacionesTarimas), "Validaciones tarimas");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cierresParciales), "Cierres parciales");

    const fecha = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, nombreArchivoSeguro(`BD_SOP_WMS_${fecha}.xlsx`));
  } catch (error) {
    console.error(error);
    alert("No se pudo exportar la base de datos a Excel. Revisa conexión o permisos de Supabase.");
  }
}
