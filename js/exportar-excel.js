/* ==========================================================================
   exportar-excel.js - "Convertir a Excel"
   Genera un .xlsx ordenado (secciones, tablas con bordes, celdas combinadas),
   con la misma estructura del formato oficial de la Orden de Trabajo.
   Requiere: xlsx-mini.js, datos.js y las funciones/variables de index.html
   ========================================================================== */

function fmtFecha(iso) {
    if (!iso) return '';
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
    return m ? `${m[3]}/${m[2]}/${m[1]}` : iso;
}

// Lee todos los datos de la orden desde el formulario
function obtenerDatosOrden() {
    const v = id => { const el = document.getElementById(id); return el ? (el.value || '').trim() : ''; };
    const selTexto = id => { const el = document.getElementById(id); return el && el.selectedIndex >= 0 && el.value ? el.options[el.selectedIndex].text : ''; };
    const tec = personalTelecom[v('selectTecnico')] || null;
    const jm = personalTelecom[v('nombreJefe')] || null;
    const ja = personalTelecom[v('nombreJefeArea')] || null;
    const esOtroCliente = v('selectCliente') === '__otro__';

    return {
        numOrden: v('numOrden'), folio: v('folioSolicitud'),
        tipoMantenimiento: v('tipoMantenimiento'), tipoArea: v('tipoArea'), medio: v('medioTransmision'),
        clasificacion: v('clasificacionMantenimiento'), estado: v('estadoOrden'),
        fechaEmision: fmtFecha(v('fechaEmision')), fechaProgramacion: fmtFecha(v('fechaProgramacion')),
        internoExterno: v('internoExterno'),
        cliente: esOtroCliente ? v('clienteManual') : (clientesBD[v('selectCliente')] ? clientesBD[v('selectCliente')].nombre : ''),
        cliNit: v('cliNit'), cliContrato: v('cliContrato'), cliServicio: v('cliServicio'), cliDireccion: v('cliDireccion'),
        cliBarrio: v('cliBarrio'), cliMunicipio: v('cliMunicipio'), cliContacto: v('cliContacto'),
        cliTelefono: v('cliTelefono'), cliCorreo: v('cliCorreo'),
        tipoRed: v('tipoRed'), medioEspecificacion: v('medioEspecificacion'),
        lugar: v('lugarTrabajo'), elemento: v('elementoIntervenido'), proveedor: v('proveedorExterno'),
        tecnico: v('nombreTecnico'), cargoTec: v('cargoTecnico'), cedulaTec: v('cedulaTecnico'),
        celularTec: v('celularTecnico'), especialidadTec: v('especialidadTecnico'), empresaTec: v('empresaTecnico'),
        motivo: v('motivoSolicitud'), diagnostico: v('diagnosticoMantenimiento'), actividades: v('actividadesRealizadas'),
        mediciones: medicionesLista.filter(m => m.parametro || m.antes || m.despues || m.observacion),
        recursos: recursosSeleccionados,
        epi: [
            ['Antecedentes de la red', v('epiAntecedentes')],
            ['Hallazgos (inspección y mediciones)', v('epiHallazgos')],
            ['Tratamiento (intervención realizada)', v('epiTratamiento')],
            ['Evolución y resultado', v('epiEvolucion')],
            ['Condición de egreso', v('epiEgreso')],
            ['Recomendaciones (plan preventivo)', v('epiRecomendaciones')]
        ],
        firmaTec: { nombre: v('nombreTecnico'), cargo: v('cargoTecnico') },
        firmaJefeMant: { nombre: jm ? jm.nombre : '', cargo: jm ? jm.cargo : '' },
        firmaJefeArea: { nombre: ja ? ja.nombre : '', cargo: ja ? ja.cargo : '' },
        firmaCliente: { nombre: v('nombreCliente') },
        fechaTermino: fmtFecha(v('fechaTermino'))
    };
}

function exportarExcel() {
    const d = obtenerDatosOrden();
    const libro = new LibroXLSX('Orden de Trabajo');

    // Anchos de columna (A..H) - total ~135 caracteres
    const ANCHOS = [13, 12, 22, 18, 14, 14, 22, 20];
    ANCHOS.forEach((w, i) => libro.ancho(i + 1, w));
    const anchoRango = (c1, c2) => ANCHOS.slice(c1 - 1, c2).reduce((a, b) => a + b, 0);
    const altoTexto = (txt, c1, c2, minimo) => {
        const cpl = anchoRango(c1, c2) * 1.05;
        const lineas = String(txt || '').split('\n').reduce((a, p) => a + Math.max(1, Math.ceil(p.length / cpl)), 0);
        return Math.max(minimo || 20, lineas * 13.5 + 6);
    };

    // ---- Estilos ----
    const VERDE = '14532D', VERDE_CLARO = 'DCFCE7', GRIS = 'F1F5F9';
    const E = {
        inst:     { b: 1, sz: 13, color: 'FFFFFF', fill: VERDE, h: 'center', v: 'center', wrap: 1, borde: VERDE },
        titulo:   { b: 1, sz: 13, color: VERDE, h: 'center', v: 'center', wrap: 1, borde: true },
        subtit:   { sz: 10, color: '475569', h: 'center', v: 'center', wrap: 1, borde: true },
        codigo:   { b: 1, sz: 10, h: 'center', v: 'center', wrap: 1, borde: true, fill: GRIS },
        seccion:  { b: 1, sz: 11, color: 'FFFFFF', fill: VERDE, h: 'left', v: 'center', borde: VERDE },
        etiqueta: { b: 1, sz: 10, color: VERDE, fill: VERDE_CLARO, h: 'left', v: 'center', wrap: 1, borde: true },
        valor:    { sz: 10, h: 'left', v: 'center', wrap: 1, borde: true },
        valorTop: { sz: 10, h: 'left', v: 'top', wrap: 1, borde: true },
        thead:    { b: 1, sz: 10, fill: GRIS, h: 'center', v: 'center', wrap: 1, borde: true },
        celda:    { sz: 10, h: 'left', v: 'center', wrap: 1, borde: true },
        celdaC:   { sz: 10, h: 'center', v: 'center', wrap: 1, borde: true },
        celdaB:   { b: 1, sz: 10, h: 'center', v: 'center', wrap: 1, borde: true },
        firmaTit: { b: 1, sz: 10, color: VERDE, fill: VERDE_CLARO, h: 'center', v: 'center', wrap: 1, borde: true },
        firmaBox: { sz: 9, color: '94A3B8', h: 'left', v: 'top', borde: true },
        firmaNom: { b: 1, sz: 10, h: 'center', v: 'center', wrap: 1, borde: true },
        firmaCar: { sz: 9, color: '475569', h: 'center', v: 'center', wrap: 1, borde: true },
        nota:     { i: 1, sz: 9, color: '64748B', h: 'center', v: 'center', wrap: 1 }
    };
    const chk = c => c ? '☒' : '☐';

    let r = 1;
    const filaVacia = (alto) => { libro.alto(r, alto || 6); r++; };
    const seccion = (txt) => { libro.poner(r, 1, r, 8, txt, E.seccion); libro.alto(r, 20); r++; };
    // Dos pares etiqueta/valor por fila: A:B etiqueta, C:D valor, E:F etiqueta, G:H valor
    const par = (l1, v1, l2, v2) => {
        libro.poner(r, 1, r, 2, l1, E.etiqueta); libro.poner(r, 3, r, 4, v1, E.valor);
        libro.poner(r, 5, r, 6, l2, E.etiqueta); libro.poner(r, 7, r, 8, v2, E.valor);
        libro.alto(r, Math.max(altoTexto(v1, 3, 4, 20), altoTexto(v2, 7, 8, 20), altoTexto(l1, 1, 2, 20), altoTexto(l2, 5, 6, 20)));
        r++;
    };
    // Campo de texto largo: etiqueta a todo el ancho + texto debajo
    const bloqueTexto = (etiqueta, texto) => {
        libro.poner(r, 1, r, 8, etiqueta, E.etiqueta); libro.alto(r, 19); r++;
        libro.poner(r, 1, r, 8, texto || '', E.valorTop); libro.alto(r, altoTexto(texto, 1, 8, 30)); r++;
    };

    // ===== ENCABEZADO =====
    libro.poner(r, 1, r + 2, 2, 'NOVA NET\nNOVA NET S.A.S.\nNIT: 901.345.678-1', E.inst);
    libro.poner(r, 3, r + 1, 6, 'ORDEN DE TRABAJO DE MANTENIMIENTO DE TELECOMUNICACIONES', E.titulo);
    libro.poner(r + 2, 3, r + 2, 6, 'Formato para Orden de Trabajo de Mantenimiento', E.subtit);
    libro.poner(r, 7, r, 8, 'Código: NOVANET-FO-OT-001', E.codigo);
    libro.poner(r + 1, 7, r + 1, 8, 'Versión: 03 | Vigencia: 2026', E.codigo);
    libro.poner(r + 2, 7, r + 2, 8, 'Estado: ' + (d.estado || '—'), E.codigo);
    libro.alto(r, 24); libro.alto(r + 1, 24); libro.alto(r + 2, 24);
    r += 3;
    filaVacia();

    // ===== I. INFORMACIÓN GENERAL =====
    seccion('SECCIÓN I - INFORMACIÓN GENERAL DE LA ORDEN');
    par('Folio de solicitud correspondiente', d.folio, 'Número de control', d.numOrden);
    par('Fecha de emisión', d.fechaEmision, 'Fecha de programación', d.fechaProgramacion);
    par('Tipo de mantenimiento',
        `${chk(d.tipoMantenimiento === 'Preventivo')} Preventivo     ${chk(d.tipoMantenimiento === 'Correctivo')} Correctivo`,
        'Alcance',
        `${chk(d.clasificacion === 'Menor')} Menor     ${chk(d.clasificacion === 'Mayor')} Mayor`);
    par('Tipo de área / infraestructura', d.tipoArea, 'Medio de transmisión', d.medio);
    par('Estado de la orden', d.estado,
        'Mantenimiento (según el técnico)',
        `${chk(d.internoExterno === 'Interno')} Interno     ${chk(d.internoExterno === 'Externo')} Externo`);
    filaVacia();

    // ===== I-A. CLIENTE =====
    seccion('SECCIÓN I-A - DATOS DEL CLIENTE');
    par('Cliente', d.cliente, 'NIT / Cédula', d.cliNit);
    par('N.º de contrato', d.cliContrato, 'Tipo de servicio', d.cliServicio);
    par('Dirección', d.cliDireccion, 'Barrio / sector', d.cliBarrio);
    par('Municipio', d.cliMunicipio, 'Persona de contacto', d.cliContacto);
    par('Teléfono de contacto', d.cliTelefono, 'Correo electrónico', d.cliCorreo);
    filaVacia();

    // ===== I-B. RED =====
    seccion('SECCIÓN I-B - DATOS DE LA RED');
    par('Tipo de red', d.tipoRed, 'Medio / especificación', d.medioEspecificacion);
    par('Ubicación / ruta / nodo (lugar del trabajo)', d.lugar, 'Elemento o equipo intervenido', d.elemento);
    filaVacia();

    // ===== PERSONAL =====
    seccion('PERSONAL ASIGNADO');
    par('Técnico asignado', d.tecnico, 'Cargo', d.cargoTec);
    par('Cédula', d.cedulaTec, 'Teléfono', d.celularTec);
    par('Especialidad', d.especialidadTec, 'Empresa / proveedor', d.empresaTec);
    libro.poner(r, 1, r, 2, 'Proveedor externo (si aplica)', E.etiqueta); libro.poner(r, 3, r, 8, d.proveedor, E.valor);
    libro.alto(r, 22); r++;
    filaVacia();

    // ===== II. MOTIVO Y DIAGNÓSTICO =====
    seccion('SECCIÓN II - UBICACIÓN Y DIAGNÓSTICO DEL EVENTO');
    bloqueTexto('Motivo de la solicitud / falla reportada', d.motivo);
    bloqueTexto('Diagnóstico del mantenimiento', d.diagnostico);
    filaVacia();

    // ===== III. ACTIVIDADES / MEDICIONES / MATERIALES =====
    seccion('SECCIÓN III - ACTIVIDADES REALIZADAS Y RECURSOS');
    bloqueTexto('Actividades realizadas', d.actividades);
    filaVacia(8);

    libro.poner(r, 1, r, 8, 'Mediciones antes y después de la intervención', E.etiqueta); libro.alto(r, 19); r++;
    libro.poner(r, 1, r, 3, 'Parámetro', E.thead); libro.poner(r, 4, r, 4, 'Unidad', E.thead);
    libro.poner(r, 5, r, 5, 'Antes', E.thead); libro.poner(r, 6, r, 6, 'Después', E.thead);
    libro.poner(r, 7, r, 8, 'Observación', E.thead); libro.alto(r, 20); r++;
    if (d.mediciones.length === 0) {
        libro.poner(r, 1, r, 8, 'Sin mediciones registradas', E.nota); r++;
    }
    d.mediciones.forEach(m => {
        libro.poner(r, 1, r, 3, m.parametro, E.celda); libro.poner(r, 4, r, 4, m.unidad, E.celdaC);
        const num = x => (x !== '' && !isNaN(Number(String(x).replace(',', '.')))) ? Number(String(x).replace(',', '.')) : x;
        libro.poner(r, 5, r, 5, num(m.antes), E.celdaC); libro.poner(r, 6, r, 6, num(m.despues), E.celdaC);
        libro.poner(r, 7, r, 8, m.observacion, E.celda);
        libro.alto(r, Math.max(20, altoTexto(m.parametro, 1, 3), altoTexto(m.observacion, 7, 8))); r++;
    });
    filaVacia(8);

    libro.poner(r, 1, r, 8, 'Cantidad de materiales y herramientas utilizados', E.etiqueta); libro.alto(r, 19); r++;
    libro.poner(r, 1, r, 1, 'Cantidad', E.thead); libro.poner(r, 2, r, 2, 'Unidad', E.thead);
    libro.poner(r, 3, r, 3, 'Tipo', E.thead); libro.poner(r, 4, r, 6, 'Descripción del material', E.thead);
    libro.poner(r, 7, r, 7, 'Origen / ubicación', E.thead); libro.poner(r, 8, r, 8, 'Observaciones', E.thead);
    libro.alto(r, 20); r++;
    if (d.recursos.length === 0) {
        libro.poner(r, 1, r, 8, 'Sin materiales ni herramientas registrados', E.nota); r++;
    }
    d.recursos.forEach(x => {
        libro.poner(r, 1, r, 1, Number(x.cantidad), E.celdaB); libro.poner(r, 2, r, 2, x.unidad, E.celdaC);
        libro.poner(r, 3, r, 3, x.tipo, E.celda); libro.poner(r, 4, r, 6, x.nombre, E.celda);
        libro.poner(r, 7, r, 7, x.origen, E.celda); libro.poner(r, 8, r, 8, x.observacion || '', E.celda);
        libro.alto(r, Math.max(20, altoTexto(x.nombre, 4, 6), altoTexto(x.origen, 7, 7), altoTexto(x.observacion, 8, 8))); r++;
    });
    filaVacia();

    // ===== EPICRISIS =====
    seccion('SECCIÓN III-B - EPICRISIS TÉCNICA Y RECOMENDACIONES');
    d.epi.forEach(([titulo, texto]) => bloqueTexto(titulo, texto));
    filaVacia();

    // ===== IV. FIRMAS =====
    seccion('SECCIÓN IV - AUTORIZACIÓN Y FIRMAS');
    const bloquesFirma = [
        { col: 1, titulo: 'Técnico ejecutante', f: d.firmaTec },
        { col: 3, titulo: 'Autorización\n(Jefe de Mantenimiento)', f: d.firmaJefeMant },
        { col: 5, titulo: 'Liberación\n(Jefe de Área)', f: d.firmaJefeArea },
        { col: 7, titulo: 'Cliente recibe', f: d.firmaCliente }
    ];
    bloquesFirma.forEach(b => libro.poner(r, b.col, r, b.col + 1, b.titulo, E.firmaTit)); libro.alto(r, 30); r++;
    bloquesFirma.forEach(b => libro.poner(r, b.col, r, b.col + 1, 'Firma:', E.firmaBox)); libro.alto(r, 55); r++;
    bloquesFirma.forEach(b => libro.poner(r, b.col, r, b.col + 1, b.f.nombre || '', E.firmaNom)); libro.alto(r, 22); r++;
    bloquesFirma.forEach(b => libro.poner(r, b.col, r, b.col + 1, b.f.cargo || '', E.firmaCar)); libro.alto(r, 22); r++;
    libro.poner(r, 1, r, 2, 'Fecha de término', E.etiqueta); libro.poner(r, 3, r, 8, d.fechaTermino, E.valor); libro.alto(r, 22); r++;
    filaVacia(10);
    libro.poner(r, 1, r, 8, 'NOVA NET S.A.S. - Sistema Integrado de Gestión de Órdenes de Mantenimiento', E.nota); r++;

    const nombre = (d.numOrden || 'Orden_de_Trabajo').replace(/[^\w\-]+/g, '_') + '_Orden_Trabajo.xlsx';
    libro.descargar(nombre);

    mostrarModal("Excel Generado", "Se descargó el archivo .xlsx con el formato ordenado de la orden de trabajo: secciones, tablas de mediciones, materiales y firmas.", "fa-file-excel", "text-teal-600 bg-teal-100");
}
