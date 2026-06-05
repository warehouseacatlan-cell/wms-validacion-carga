let datosExcel = [];
let pedidoSeleccionado = [];
let tarimasValidadas = [];
let cierreParcial = null;
let lectorCamara = null;
let escanerActivo = false;
let modoEscaneoCamara = "auto";

function $(id) {
  return document.getElementById(id);
}

function mostrarSeccion(id) {
  document.querySelectorAll(".seccion").forEach(sec => {
    sec.style.display = "none";
  });

  const seccion = $(id);

  if (seccion) {
    seccion.style.display = "block";
  }
}

function leerExcel() {
  const archivo = $("archivoExcel")?.files[0];

  if (!archivo) {
    alert("Seleccione un archivo Excel");
    return;
  }

  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(hoja, { header: 1 });

      procesarExcel(filas);
    } catch (error) {
      console.error(error);
      alert("No se pudo leer el Excel. Revise el archivo.");
    }
  };

  reader.readAsArrayBuffer(archivo);
}

function procesarExcel(filas) {
  datosExcel = [];
  pedidoSeleccionado = [];
  tarimasValidadas = [];
  cierreParcial = null;

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
  mostrarPedidosDetectados();
  guardarEstado();
}

function convertirCantidad(valor) {
  if (typeof valor === "number") return valor;

  return Number(
    String(valor)
      .replace(/,/g, "")
      .replace(/\s/g, "")
  ) || 0;
}

function normalizarSKU(valor) {
  return String(valor || "")
    .trim()
    .replace(/^\[/, "")
    .replace(/\]$/, "")
    .toUpperCase();
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
  const datos = {
    sku: "",
    lote: "",
    caducidad: "",
    cantidad: ""
  };

  if (!textoOriginal) return datos;

  const texto = textoOriginal
    .replace(/\r/g, "\n")
    .replace(/\|/g, "\n")
    .replace(/;/g, "\n")
    .replace(/,/g, "\n");

  // 1) QR en formato JSON: {"sku":"ABC", "lote":"L1", "caducidad":"2026-05-30", "cantidad":20}
  try {
    const json = JSON.parse(textoOriginal);
    datos.sku = json.sku || json.SKU || json.codigo || json.codigoProducto || json.producto || "";
    datos.lote = json.lote || json.LOTE || json.batch || json.Batch || "";
    datos.caducidad = normalizarFecha(json.caducidad || json.CADUCIDAD || json.exp || json.EXP || json.fechaCaducidad || "");
    datos.cantidad = json.cantidad || json.CANTIDAD || json.qty || json.QTY || json.piezas || "";
    return datos;
  } catch (_) {}

  // 2) QR con claves: SKU=ABC; LOTE=L1; CAD=30/05/2026; CANT=20
  datos.sku = extraerValorPorClaves(texto, ["SKU", "CODIGO", "CÓDIGO", "PRODUCTO", "ITEM", "ARTICULO", "ARTÍCULO", "CLAVE"]);
  datos.lote = extraerValorPorClaves(texto, ["LOTE", "LOT", "BATCH"]);
  datos.caducidad = normalizarFecha(extraerValorPorClaves(texto, ["CADUCIDAD", "CAD", "EXP", "VENCE", "VENCIMIENTO", "FECHA CADUCIDAD"]));
  datos.cantidad = extraerValorPorClaves(texto, ["CANTIDAD", "CANT", "QTY", "PIEZAS", "PZAS", "PZA", "PCS"]);

  // 3) Formato GS1 visible con paréntesis: (240)SKU(10)LOTE(17)260530(30)20
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

  // 4) Texto tipo producto de Excel: [SKU] Descripción
  if (!datos.sku) {
    const entreCorchetes = textoOriginal.match(/\[(.*?)\]/);
    if (entreCorchetes) datos.sku = entreCorchetes[1];
  }

  // 5) Si no trae claves y parece un SKU simple, usar todo como SKU
  if (!datos.sku && /^[A-Za-z0-9._-]{3,40}$/.test(textoOriginal)) {
    datos.sku = textoOriginal;
  }

  datos.sku = normalizarSKU(datos.sku);
  datos.lote = String(datos.lote || "").trim();
  datos.cantidad = datos.cantidad ? String(datos.cantidad).replace(/[^0-9.]/g, "") : "";

  return datos;
}

function extraerValorPorClaves(texto, claves) {
  for (const clave of claves) {
    const claveSegura = clave.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const patron = new RegExp(`(?:^|\\n|\\s)${claveSegura}\\s*[:=\\-#]\\s*([^\\n]+)`, "i");
    const match = texto.match(patron);
    if (match && match[1]) {
      return match[1].trim();
    }
  }

  return "";
}


function normalizarFecha(valor) {
  const texto = String(valor || "").trim();
  if (!texto) return "";

  // yyyy-mm-dd
  let match = texto.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return `${match[1]}-${match[2]}-${match[3]}`;

  // dd/mm/yyyy o dd-mm-yyyy
  match = texto.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (match) {
    const dd = match[1].padStart(2, "0");
    const mm = match[2].padStart(2, "0");
    const yyyy = match[3];
    return `${yyyy}-${mm}-${dd}`;
  }

  // yymmdd para GS1
  match = texto.match(/^(\d{2})(\d{2})(\d{2})$/);
  if (match) return convertirFechaGS1(texto);

  return "";
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

  if ($("datosLeidosQR")) {
    $("datosLeidosQR").value = texto;
  }

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
  if (datos.caducidad) $("caducidadEscaneada").value = datos.caducidad;
  if (datos.cantidad) $("cantidadTarima").value = datos.cantidad;

  if (!datos.sku && texto) {
    $("skuEscaneado").value = extraerSKUDesdeScan(texto);
  }

  if (!$("loteEscaneado").value) {
    $("loteEscaneado").focus();
  } else if (!$("caducidadEscaneada").value) {
    $("caducidadEscaneada").focus();
  } else if (!$("cantidadTarima").value) {
    $("cantidadTarima").focus();
  } else {
    $("fotoTarima1").focus();
  }
}

function enfocarCampoDespuesDeSKU() {
  if (!$("loteEscaneado").value) {
    $("loteEscaneado").focus();
  } else if (!$("caducidadEscaneada").value) {
    $("caducidadEscaneada").focus();
  } else if (!$("cantidadTarima").value) {
    $("cantidadTarima").focus();
  } else {
    $("fotoTarima1").focus();
  }
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
  const pedidosUnicos = [...new Set(datosExcel.map(x => x.pedido))];

  $("resumenExcel").innerHTML = `
    <h4>Excel cargado correctamente</h4>
    <p>SKU consolidados: <b>${datosExcel.length}</b></p>
    <p>Pedidos encontrados: <b>${pedidosUnicos.length}</b></p>
  `;

  let opciones = `<option value="">Seleccione un pedido</option>`;

  pedidosUnicos.forEach(pedido => {
    opciones += `<option value="${escaparHTML(pedido)}">${escaparHTML(pedido)}</option>`;
  });

  $("selectorPedido").innerHTML = `
    <h4>Seleccione el pedido a validar</h4>

    <select id="pedidoDetectado">
      ${opciones}
    </select>

    <button onclick="seleccionarPedido()">Cargar pedido seleccionado</button>
    <button onclick="verDashboardPedidos()">Ver dashboard de pedidos</button>
    <button onclick="verHistorial()">Ver historial</button>
  `;
}

function seleccionarPedido() {
  const pedido = $("pedidoDetectado")?.value;

  if (!pedido) {
    alert("Seleccione un pedido");
    return;
  }

  cargarPedidoDesdeDashboard(pedido);
}

function cargarPedidoDesdeDashboard(pedido) {
  pedidoSeleccionado = datosExcel.filter(x => x.pedido === pedido);

  if (pedidoSeleccionado.length === 0) {
    alert("No se encontró información para este pedido");
    return;
  }

  if (cierreParcial && cierreParcial.pedido !== pedido) {
    cierreParcial = null;
  }

  $("pedido").value = pedidoSeleccionado[0].pedido;
  $("cliente").value = pedidoSeleccionado[0].cliente;

  actualizarAvance();
  guardarEstado();

  mostrarSeccion("pasoDatos");
}

function validarDatosGenerales() {
  const pedido = $("pedido").value.trim();
  const cliente = $("cliente").value.trim();
  const chofer = $("chofer").value.trim();
  const validador = $("validador").value.trim();

  if (!pedido || !cliente || !chofer || !validador) {
    alert("Debe completar Pedido, Cliente, Chofer y Validador");
    return;
  }

  actualizarAvance();
  guardarEstado();
  guardarHistorialLigero(obtenerEstatusPedido(pedido));

  mostrarSeccion("pasoValidacion");
  enfocarSKU();
}

async function guardarTarima() {
  const sku = extraerSKUDesdeScan($("skuEscaneado").value);
  const lote = $("loteEscaneado").value.trim();
  const caducidad = $("caducidadEscaneada").value;
  const cantidad = Number($("cantidadTarima").value);

  const foto1 = $("fotoTarima1").files[0];
  const foto2 = $("fotoTarima2").files[0];

  if (!sku || !lote || !caducidad || !cantidad) {
    alert("Debe completar SKU, lote, caducidad y cantidad");
    return;
  }

  if (cantidad <= 0) {
    alert("La cantidad debe ser mayor a cero");
    return;
  }

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

  const nuevoTotal = lineaPedido.cantidadValidada + cantidad;

  if (nuevoTotal > lineaPedido.cantidadPedida) {
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

  lineaPedido.cantidadValidada = nuevoTotal;

  const lineaExcel = datosExcel.find(x =>
    x.pedido === lineaPedido.pedido &&
    normalizarSKU(x.sku) === normalizarSKU(lineaPedido.sku)
  );

  if (lineaExcel) {
    lineaExcel.cantidadValidada = lineaPedido.cantidadValidada;
  }

  tarimasValidadas.push({
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

  limpiarFormularioTarima();
  actualizarAvance();

  const estatus = pedidoCompletado() ? "COMPLETADO" : "EN PROCESO";

  guardarEstado();
  guardarHistorialLigero(estatus);

  if (pedidoCompletado()) {
    alert("PEDIDO COMPLETADO CORRECTAMENTE");
    mostrarSeccion("pasoAvance");
  } else {
    alert("Tarima guardada correctamente");
  }
}

function limpiarFormularioTarima() {
  $("skuEscaneado").value = "";
  $("loteEscaneado").value = "";
  $("caducidadEscaneada").value = "";
  $("cantidadTarima").value = "";
  $("fotoTarima1").value = "";
  $("fotoTarima2").value = "";
  if ($("datosLeidosQR")) $("datosLeidosQR").value = "";

  enfocarSKU();
}

function limpiarCapturaTarima() {
  $("skuEscaneado").value = "";
  $("loteEscaneado").value = "";
  $("caducidadEscaneada").value = "";
  $("cantidadTarima").value = "";
  if ($("datosLeidosQR")) $("datosLeidosQR").value = "";
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
    totalPedido += item.cantidadPedida;
    totalValidado += item.cantidadValidada;

    const pendiente = item.cantidadPedida - item.cantidadValidada;

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

  const porcentaje = totalPedido > 0
    ? Math.round((totalValidado / totalPedido) * 100)
    : 0;

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
    pedidoSeleccionado.every(item => item.cantidadValidada === item.cantidadPedida);
}

function obtenerEstatusPedido(pedido) {
  if (cierreParcial && cierreParcial.pedido === pedido) {
    return "COMPLETADO PARCIAL";
  }

  const lineas = datosExcel.filter(x => x.pedido === pedido);

  if (lineas.length === 0) return "PENDIENTE";

  const totalPedido = lineas.reduce((suma, item) => suma + item.cantidadPedida, 0);
  const totalValidado = lineas.reduce((suma, item) => suma + item.cantidadValidada, 0);

  if (totalPedido > 0 && totalValidado === totalPedido) return "COMPLETADO";
  if (totalValidado > 0) return "EN PROCESO";
  return "PENDIENTE";
}

function cerrarPedidoParcial() {
  if (pedidoSeleccionado.length === 0) {
    alert("Primero seleccione un pedido");
    return;
  }

  const pedido = $("pedido").value.trim();

  if (pedidoCompletado()) {
    alert("El pedido ya está completo. No requiere cierre parcial.");
    return;
  }

  const totalValidado = pedidoSeleccionado.reduce((suma, item) => suma + item.cantidadValidada, 0);

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

  cierreParcial = {
    pedido,
    cliente: $("cliente").value.trim(),
    motivo: motivo.trim(),
    autorizo: autorizo.trim(),
    comentario: comentario.trim(),
    fechaHora: new Date().toLocaleString()
  };

  actualizarAvance();
  guardarEstado();
  guardarHistorialLigero("COMPLETADO PARCIAL");

  alert("Pedido cerrado como COMPLETADO PARCIAL");
  mostrarSeccion("pasoAvance");
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

function verDashboardPedidos() {
  if (datosExcel.length === 0) {
    alert("Primero cargue un Excel");
    return;
  }

  const pedidosUnicos = [...new Set(datosExcel.map(x => x.pedido))];

  let completados = 0;
  let parciales = 0;
  let enProceso = 0;
  let pendientes = 0;
  let html = "";

  pedidosUnicos.forEach(pedido => {
    const lineas = datosExcel.filter(x => x.pedido === pedido);
    const cliente = lineas[0]?.cliente || "";

    const totalPedido = lineas.reduce((suma, item) => suma + item.cantidadPedida, 0);
    const totalValidado = lineas.reduce((suma, item) => suma + item.cantidadValidada, 0);
    const pendiente = totalPedido - totalValidado;
    const estatus = obtenerEstatusPedido(pedido);

    if (estatus === "COMPLETADO") completados++;
    else if (estatus === "COMPLETADO PARCIAL") parciales++;
    else if (estatus === "EN PROCESO") enProceso++;
    else pendientes++;

    html += `
      <div class="linea-avance">
        <b>Pedido:</b> ${escaparHTML(pedido)}<br>
        <b>Cliente:</b> ${escaparHTML(cliente)}<br>
        <b>Total pedido:</b> ${totalPedido}<br>
        <b>Total validado:</b> ${totalValidado}<br>
        <b>Pendiente:</b> ${pendiente}<br>
        <b>Estatus:</b> ${estatus}<br>
        <button onclick="cargarPedidoDesdeDashboard('${escaparAtributo(pedido)}')">Abrir pedido</button>
        <button onclick="generarPDFDesdePedido('${escaparAtributo(pedido)}')">PDF</button>
      </div>
    `;
  });

  const resumen = `
    <div class="resultado">
      <h4>Resumen general</h4>
      <p>Pedidos encontrados: <b>${pedidosUnicos.length}</b></p>
      <p>Completados: <b>${completados}</b></p>
      <p>Completados parciales: <b>${parciales}</b></p>
      <p>En proceso: <b>${enProceso}</b></p>
      <p>Pendientes: <b>${pendientes}</b></p>
    </div>
  `;

  $("dashboardPedidos").innerHTML = resumen + html;
  mostrarSeccion("pasoDashboard");
}

function guardarEstado() {
  const estado = {
    datosExcel,
    pedidoSeleccionado,
    tarimasValidadas,
    cierreParcial,
    pedido: $("pedido")?.value || "",
    cliente: $("cliente")?.value || "",
    chofer: $("chofer")?.value || "",
    validador: $("validador")?.value || "",
    fechaGuardado: new Date().toLocaleString()
  };

  try {
    localStorage.setItem("sopwms_estado", JSON.stringify(estado));
  } catch (error) {
    console.error(error);
    alert("No se pudo guardar en memoria. Puede que las fotos sean muy pesadas.");
  }
}

function recuperarEstado() {
  const estadoGuardado = localStorage.getItem("sopwms_estado");

  if (!estadoGuardado) return;

  let estado;

  try {
    estado = JSON.parse(estadoGuardado);
  } catch (error) {
    localStorage.removeItem("sopwms_estado");
    return;
  }

  const tieneExcel = estado.datosExcel && estado.datosExcel.length > 0;
  const tienePedido = estado.pedidoSeleccionado && estado.pedidoSeleccionado.length > 0;
  const tieneTarimas = estado.tarimasValidadas && estado.tarimasValidadas.length > 0;

  if (!tieneExcel && !tienePedido && !tieneTarimas) {
    localStorage.removeItem("sopwms_estado");
    return;
  }

  datosExcel = estado.datosExcel || [];
  pedidoSeleccionado = estado.pedidoSeleccionado || [];
  tarimasValidadas = estado.tarimasValidadas || [];
  cierreParcial = estado.cierreParcial || null;

  $("pedido").value = estado.pedido || "";
  $("cliente").value = estado.cliente || "";
  $("chofer").value = estado.chofer || "";
  $("validador").value = estado.validador || "";

  if (datosExcel.length > 0) {
    mostrarPedidosDetectados();
  }

  if (!tienePedido && !tieneTarimas) {
    mostrarSeccion("pasoExcel");
    return;
  }

  const confirmar = confirm("Se encontró una validación en proceso.\n\n¿Deseas continuar?");

  if (!confirmar) {
    localStorage.removeItem("sopwms_estado");
    location.reload();
    return;
  }

  actualizarAvance();
  mostrarSeccion("pasoValidacion");
  enfocarSKU();
}

function limpiarMemoriaPedido() {
  const confirmar = confirm("¿Deseas borrar la validación guardada de este equipo?");

  if (!confirmar) return;

  localStorage.removeItem("sopwms_estado");

  alert("Memoria del pedido eliminada correctamente");
  location.reload();
}

function borrarTodoLocal() {
  const confirmar = confirm("Se borrará toda la información almacenada en este navegador.");

  if (!confirmar) return;

  localStorage.removeItem("sopwms_estado");
  localStorage.removeItem("sopwms_historial");

  alert("Memoria local eliminada correctamente");
  location.reload();
}

function obtenerSemana(fecha) {
  const f = new Date(fecha);
  const primerDia = new Date(f.getFullYear(), 0, 1);
  const dias = Math.floor((f - primerDia) / (24 * 60 * 60 * 1000));
  const semana = Math.ceil((dias + primerDia.getDay() + 1) / 7);

  return `${f.getFullYear()}-W${semana}`;
}

function guardarHistorialLigero(estatusFinal) {
  if (pedidoSeleccionado.length === 0) return;

  const hoy = new Date();

  const totalPedido = pedidoSeleccionado.reduce((suma, item) => suma + item.cantidadPedida, 0);
  const totalValidado = pedidoSeleccionado.reduce((suma, item) => suma + item.cantidadValidada, 0);

  const registro = {
    pedido: $("pedido").value,
    cliente: $("cliente").value,
    chofer: $("chofer").value,
    validador: $("validador").value,
    fecha: hoy.toISOString().slice(0, 10),
    mes: hoy.toISOString().slice(0, 7),
    semana: obtenerSemana(hoy),
    fechaHora: hoy.toLocaleString(),
    estatus: estatusFinal,
    totalPedido,
    totalValidado,
    pendiente: totalPedido - totalValidado,
    pdfGenerado: false,
    cierreParcial: cierreParcial && cierreParcial.pedido === $("pedido").value ? cierreParcial : null,
    lineasPedido: JSON.parse(JSON.stringify(pedidoSeleccionado)),
    tarimasValidadas: JSON.parse(JSON.stringify(tarimasValidadas.filter(t => t.pedido === $("pedido").value)))
  };

  let historial = JSON.parse(localStorage.getItem("sopwms_historial")) || [];

  const indice = historial.findIndex(item => item.pedido === registro.pedido);

  if (indice >= 0) {
    historial[indice] = registro;
  } else {
    historial.push(registro);
  }

  try {
    localStorage.setItem("sopwms_historial", JSON.stringify(historial));
  } catch (error) {
    console.error(error);
    alert("No se pudo guardar el historial completo. Las fotos pueden estar ocupando mucha memoria.");
  }
}

function verHistorial() {
  const historial = JSON.parse(localStorage.getItem("sopwms_historial")) || [];
  const busqueda = $("buscarHistorial")?.value.toLowerCase() || "";
  const periodo = $("filtroPeriodo")?.value || "todos";

  const hoy = new Date();
  const diaActual = hoy.toISOString().slice(0, 10);
  const mesActual = hoy.toISOString().slice(0, 7);
  const semanaActual = obtenerSemana(hoy);

  const filtrado = historial.filter(item => {
    const coincideBusqueda =
      String(item.pedido || "").toLowerCase().includes(busqueda) ||
      String(item.cliente || "").toLowerCase().includes(busqueda);

    let coincidePeriodo = true;

    if (periodo === "dia") coincidePeriodo = item.fecha === diaActual;
    if (periodo === "semana") coincidePeriodo = item.semana === semanaActual;
    if (periodo === "mes") coincidePeriodo = item.mes === mesActual;

    return coincideBusqueda && coincidePeriodo;
  });

  if (filtrado.length === 0) {
    $("historialValidaciones").innerHTML = `<p>No hay registros encontrados.</p>`;
    mostrarSeccion("pasoHistorial");
    return;
  }

  let html = "";

  filtrado.forEach(item => {
    html += `
      <div class="linea-avance">
        <b>Pedido:</b> ${escaparHTML(item.pedido)}<br>
        <b>Cliente:</b> ${escaparHTML(item.cliente)}<br>
        <b>Chofer:</b> ${escaparHTML(item.chofer)}<br>
        <b>Validador:</b> ${escaparHTML(item.validador)}<br>
        <b>Fecha:</b> ${escaparHTML(item.fechaHora)}<br>
        <b>Total pedido:</b> ${item.totalPedido}<br>
        <b>Total validado:</b> ${item.totalValidado}<br>
        <b>Pendiente:</b> ${item.pendiente}<br>
        <b>Estatus:</b> ${escaparHTML(item.estatus)}<br>
        <b>PDF:</b> ${item.pdfGenerado ? "GENERADO" : "PENDIENTE"}<br><br>
        <button onclick="generarPDFDesdeHistorial('${escaparAtributo(item.pedido)}')">Generar PDF</button>
      </div>
    `;
  });

  $("historialValidaciones").innerHTML = html;
  mostrarSeccion("pasoHistorial");
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

  const estatusPDF = datos.cierreParcial
    ? "COMPLETADO PARCIAL"
    : pendiente === 0
      ? "COMPLETADO"
      : "INCOMPLETO";

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
    y = escribirLineaPDF(
      doc,
      `Pedido: ${item.cantidadPedida} | Validado: ${item.cantidadValidada} | Pendiente: ${pendienteSku}`,
      14,
      y,
      180,
      5
    );
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

  marcarPDFGenerado(datos.pedido);
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
  return String(texto || "")
    .replace(/[^a-zA-Z0-9_-]/g, "_")
    .substring(0, 40);
}

function marcarPDFGenerado(pedido) {
  let historial = JSON.parse(localStorage.getItem("sopwms_historial")) || [];

  historial = historial.map(item => {
    if (item.pedido === pedido) {
      item.pdfGenerado = true;
      item.fechaPDF = new Date().toLocaleString();
    }
    return item;
  });

  localStorage.setItem("sopwms_historial", JSON.stringify(historial));

  alert("PDF generado correctamente");
}

function generarPDFDesdePedido(pedido) {
  const lineas = datosExcel.filter(x => x.pedido === pedido);
  const tarimas = tarimasValidadas.filter(x => x.pedido === pedido);

  if (lineas.length === 0) {
    alert("No se encontró información para generar PDF");
    return;
  }

  const cliente = lineas[0].cliente || "";
  const parcial = cierreParcial && cierreParcial.pedido === pedido ? cierreParcial : null;

  generarPDFPedido(
    lineas,
    tarimas,
    {
      pedido,
      cliente,
      chofer: $("chofer")?.value || "No capturado",
      validador: $("validador")?.value || "No capturado",
      cierreParcial: parcial
    }
  );
}

function generarPDFDesdeHistorial(pedido) {
  const historial = JSON.parse(localStorage.getItem("sopwms_historial")) || [];
  const registro = historial.find(x => x.pedido === pedido);

  if (!registro) {
    alert("No se encontró el pedido en historial");
    return;
  }

  const lineas = registro.lineasPedido || datosExcel.filter(x => x.pedido === pedido);
  const tarimas = registro.tarimasValidadas || tarimasValidadas.filter(x => x.pedido === pedido);

  if (lineas.length === 0) {
    alert("El historial no tiene detalle suficiente para generar PDF. Abra el pedido desde el dashboard.");
    return;
  }

  generarPDFPedido(
    lineas,
    tarimas,
    {
      pedido: registro.pedido,
      cliente: registro.cliente,
      chofer: registro.chofer,
      validador: registro.validador,
      cierreParcial: registro.cierreParcial || null
    }
  );
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
        const maxAncho = 1000;
        const escala = Math.min(1, maxAncho / img.width);
        const ancho = Math.round(img.width * escala);
        const alto = Math.round(img.height * escala);

        const canvas = document.createElement("canvas");
        canvas.width = ancho;
        canvas.height = alto;

        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, ancho, alto);

        resolve(canvas.toDataURL("image/jpeg", 0.72));
      };

      img.onerror = reject;
      img.src = evento.target.result;
    };

    reader.onerror = reject;
    reader.readAsDataURL(archivo);
  });
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

  const config = {
    fps: 10,
    qrbox: { width: 280, height: 180 }
  };

  if (formatos) config.formatsToSupport = formatos;

  lectorCamara.start(
    { facingMode: "environment" },
    config,
    codigo => {
      procesarCodigoCamara(codigo);
    },
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

document.addEventListener("DOMContentLoaded", () => {
  recuperarEstado();

  const skuInput = $("skuEscaneado");
  const loteInput = $("loteEscaneado");

  if (skuInput) {
    skuInput.addEventListener("keydown", event => {
      if (event.key === "Enter") {
        event.preventDefault();
        aplicarDatosEscaneados(skuInput.value, "auto");
      }
    });

    skuInput.addEventListener("change", () => {
      const texto = skuInput.value || "";
      if (texto.includes("=") || texto.includes(":" ) || texto.includes(";") || texto.includes("|") || texto.includes("{")) {
        aplicarDatosEscaneados(texto, "auto");
      } else {
        skuInput.value = extraerSKUDesdeScan(texto);
      }
    });
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
