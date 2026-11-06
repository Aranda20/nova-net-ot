/* ==========================================================================
   NOVA NET - Bases de datos y casos de la Orden de Trabajo
   Fuente: Orden_de_Trabajo-RDCEX.xlsx (hojas Clientes, Personal, Listas,
   OT Correctivo, OT Preventivo y Caso - Epicrisis)
   ========================================================================== */

// ---------------------------------------------------------------------------
// LISTA: TIPOS DE RED (hoja "Listas")
// ---------------------------------------------------------------------------
const TIPOS_RED = [
    "Red Coaxial Externa (.500-.750)",
    "Red Coaxial Interna (RG11-RG6)",
    "Fibra Óptica Interna",
    "Fibra Óptica Externa",
    "Par trenzado (UTP)",
    "Inalámbrica / radioenlace",
    "Otra"
];

// ---------------------------------------------------------------------------
// BASE DE DATOS DE INVENTARIO OFICIAL
// (se conservan los elementos originales y se agregan los de Red Coaxial Externa)
// ---------------------------------------------------------------------------
const baseDatosInventario = [
    { tipo: "Material", nombre: "Alcohol Isopropílico 99% (Frasco)", unidad: "und", origen: "Catálogo Inventario" },
    { tipo: "Material", nombre: "Paños libres de pelusa (Lint-free wipes)", unidad: "paquete", origen: "Catálogo Inventario" },
    { tipo: "Material", nombre: "Conector Mecánico Óptico SC/APC", unidad: "und", origen: "Catálogo Inventario" },
    { tipo: "Material", nombre: "Divisor Troncal 2 Vías", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Divisor Troncal 3 Vías", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Divisor Troncal 4 Vías", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Acoplador Direccional (DC)", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Multitap de 2 Salidas", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Multitap de 4 Salidas", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Multitap de 8 Salidas", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Ecualizador de Línea (LEQ)", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Insertor de Potencia en Línea (LPI)", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Dispositivo Activo", nombre: "Nodo Óptico HFC", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Dispositivo Activo", nombre: "Amplificador Troncal / Mini Troncal", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Dispositivo Activo", nombre: "Extensor de Línea (Line Extender - LE)", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Dispositivo Activo", nombre: "Fuente de Alimentación AC Ferroresonante", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Cable Coaxial", nombre: "Cable Rígido Troncal .750\" Autosoportado", unidad: "Metros", origen: "Inventario HFC" },
    { tipo: "Cable Coaxial", nombre: "Cable Rígido Distribución .500\" Autosoportado", unidad: "Metros", origen: "Inventario HFC" },
    { tipo: "Conector", nombre: "Conector Tipo Pin .500\"", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Conector", nombre: "Conector Tipo Pin .750\"", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Conector", nombre: "Manga Termocontráctil", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Herramienta", nombre: "Medidor de Potencia Óptica (Power Meter / VFL)", unidad: "und", origen: "Catálogo Herramientas" },
    { tipo: "Herramienta", nombre: "Curvador de Cable", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Herramienta", nombre: "Máquina Cosedora / Devanadora", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Herramienta", nombre: "Tensora de Cadena (Chicharra)", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Herramienta", nombre: "Herramienta Pelachaqueta", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Herramienta", nombre: "Alicate Cortacable Circular", unidad: "Und", origen: "Inventario HFC" },
    { tipo: "Herramienta", nombre: "Soplete a Gas con Boquilla No Reductora", unidad: "Und", origen: "Inventario HFC" },

    // --- NUEVOS: materiales de Red Coaxial Externa (casos del Excel) ---
    { tipo: "Conector", nombre: "Conector .500 nuevo con sello", unidad: "und", origen: "Inventario HFC" },
    { tipo: "Conector", nombre: "Conector .500 con sello (reserva)", unidad: "und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Unión recta .500", unidad: "und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Funda termoencogible", unidad: "und", origen: "Inventario HFC" },
    { tipo: "Material", nombre: "Cinta autofusionante", unidad: "rollo", origen: "Inventario HFC" },
    { tipo: "Cable Coaxial", nombre: "Cable coaxial .500 (reserva)", unidad: "m", origen: "Inventario HFC" }
];

// ---------------------------------------------------------------------------
// PERSONAL Y TÉCNICOS
// Los dos técnicos originales ("1" y "2") se conservan. Se agrega la hoja
// "Personal" del Excel (PE-001 ... PE-008).
// ---------------------------------------------------------------------------
const personalTelecom = {
    "1": { nombre: "Dairon Andrey Aranda", cargo: "Tecnólogo Telecomunicaciones", celular: "3000000001", cedula: "", especialidad: "", tipo: "Interno", empresa: "NOVA NET S.A.S." },
    "2": { nombre: "Juan Miguel Fernandez Palta", cargo: "Ingeniero Redes", celular: "3000000002", cedula: "", especialidad: "", tipo: "Interno", empresa: "NOVA NET S.A.S." },
    "PE-001": { nombre: "Carlos Mejía", cargo: "Técnico de mantenimiento", cedula: "1061000101", celular: "3205550101", especialidad: "Redes coaxiales y HFC", tipo: "Interno", empresa: "Interno" },
    "PE-002": { nombre: "Ana Restrepo", cargo: "Jefe de Mantenimiento", cedula: "1061000102", celular: "3205550102", especialidad: "Gestión de mantenimiento", tipo: "Interno", empresa: "Interno" },
    "PE-003": { nombre: "Luis Ortega", cargo: "Jefe de Operaciones", cedula: "1061000103", celular: "3205550103", especialidad: "Operación de redes", tipo: "Interno", empresa: "Interno" },
    "PE-004": { nombre: "Diana Paz", cargo: "Técnica de fibra óptica", cedula: "1061000104", celular: "3205550104", especialidad: "Fusión y medición OTDR", tipo: "Interno", empresa: "Interno" },
    "PE-005": { nombre: "Jorge Cuéllar", cargo: "Técnico de cableado estructurado", cedula: "1061000105", celular: "3205550105", especialidad: "UTP y certificación de cableado", tipo: "Interno", empresa: "Interno" },
    "PE-006": { nombre: "Andrés Rojas", cargo: "Técnico de radioenlaces", cedula: "1061000106", celular: "3205550106", especialidad: "Radioenlaces y antenas", tipo: "Externo", empresa: "TeleRedes del Sur S.A.S." },
    "PE-007": { nombre: "Marcela Ocampo", cargo: "Jefe de Atención al Cliente", cedula: "1061000107", celular: "3205550107", especialidad: "Servicio al cliente", tipo: "Interno", empresa: "Interno" },
    // En el Excel este registro está incompleto (sin cargo, especialidad ni tipo)
    "PE-008": { nombre: "Maria Fernanda Garcia", cargo: "", cedula: "10611000108", celular: "3205550108", especialidad: "", tipo: "", empresa: "" }
};

// ---------------------------------------------------------------------------
// BASE DE DATOS DE CLIENTES (hoja "Clientes")
// ---------------------------------------------------------------------------
const clientesBD = {
    "CL-001": { nombre: "Junta de Acción Comunal Barrio El Mirador", nit: "900.456.123-4", contrato: "CT-2021-0318", direccion: "Calle 14 # 8-25, salón comunal", barrio: "El Mirador", municipio: "Piendamó", contacto: "Marta Lucía Ibarra (presidenta JAC)", telefono: "3105550142", correo: "jacelmirador@example.com", servicio: "TV + Internet (red coaxial)" },
    "CL-002": { nombre: "Hotel Casa del Río S.A.S.", nit: "901.222.567-1", contrato: "CT-2022-0107", direccion: "Carrera 6 # 20-14", barrio: "Centro", municipio: "Timbío", contacto: "Andrea Valencia (administradora)", telefono: "3115550177", correo: "administracion@example.com", servicio: "Fibra óptica dedicada" },
    "CL-003": { nombre: "Institución Educativa San Rafael", nit: "891.500.348-9", contrato: "CT-2023-0045", direccion: "Carrera 3 # 5-60", barrio: "San Rafael", municipio: "Cajibío", contacto: "Hernán Díaz (coordinador de sistemas)", telefono: "3125550119", correo: "sistemas@example.com", servicio: "Red UTP (cableado estructurado)" },
    "CL-004": { nombre: "Cooperativa Café del Cauca Ltda.", nit: "800.789.456-2", contrato: "CT-2020-0290", direccion: "Km 3 vía a la vereda La Esperanza", barrio: "Vereda La Esperanza", municipio: "El Tambo", contacto: "Sandra Muñoz (jefe administrativa)", telefono: "3135550163", correo: "gerencia@example.com", servicio: "Radioenlace punto a punto" },
    "CL-005": { nombre: "Clínica Los Andes S.A.", nit: "890.321.654-7", contrato: "CT-2024-0012", direccion: "Calle 9 # 11-40", barrio: "La Pamba", municipio: "Santander de Quilichao", contacto: "Julián Torres (ingeniero de TI)", telefono: "3145550188", correo: "ti@example.com", servicio: "Fibra óptica dedicada" },
    "CL-006": { nombre: "Almacén Distrisur", nit: "900.654.987-3", contrato: "CT-2024-0136", direccion: "Carrera 7 # 15-22, local 2", barrio: "Centro", municipio: "Timbío", contacto: "Paola Ruiz (propietaria)", telefono: "3155550134", correo: "distrisur@example.com", servicio: "Internet por cable UTP" }
};

// ---------------------------------------------------------------------------
// CASO 1 - CORRECTIVO: Red Coaxial Externa (OT-2026-0087)
// Hojas "OT Correctivo" y "Caso - Epicrisis" del Excel
// ---------------------------------------------------------------------------
const CASO_CORRECTIVO = {
    numOrden: "OT-2026-0087",
    folio: "SOL-2026-0145",
    tipoMantenimiento: "Correctivo",
    tipoArea: "Aérea",
    medioTransmision: "Cable Coaxial",
    clasificacion: "Menor",
    fechaEmision: "2026-10-19",
    fechaProgramacion: "2026-10-20",
    estado: "Cerrada",
    tecnicoId: "PE-001",
    clienteId: "CL-001",
    tipoRed: "Red Coaxial Externa (.500-.750)",
    medioEspecificacion: "Cable coaxial .500 (cobre), tramo aéreo de unos 420 m",
    lugarTrabajo: "Nodo N-07, tramo aéreo entre los amplificadores A-07 y A-08 (sector 3)",
    elementoIntervenido: "Conector de salida del amplificador A-08 y tramo de cable dañado (a 60 m de A-08)",
    proveedorExterno: "No aplica (Personal Interno NOVA NET)",
    motivo: "Entre el 17 y el 19 de octubre de 2026, 23 usuarios residenciales del sector 3 reportaron pixelación y congelamiento de los canales de televisión (sobre todo en la noche) y cortes de internet en los días de lluvia. El servicio mejoraba por momentos cuando el clima estaba seco.",
    diagnostico: "Ingreso de agua por el conector de salida del amplificador A-08 (flojo, sin funda de sello, con humedad y corrosión en el conductor central y la rosca) y por un corte de unos 5 cm en el forro del cable, a 60 m de A-08, con la malla expuesta y mojada. La corrosión y la humedad aumentaron la atenuación y desadaptaron la impedancia de la línea (MER y pérdida de retorno bajos). Causa raíz: falta de mantenimiento preventivo (sellos y forro sin revisar; sin registro en los últimos dos años).",
    actividades: "1. Verificar la presencia de voltaje AC en la línea antes de intervenir.\n2. Inspeccionar el tramo A-07 – A-08 y medir en el tap más lejano (valores iniciales).\n3. Cortar y retirar el tramo de cable dañado.\n4. Instalar conectores .500 nuevos con sello y una unión recta.\n5. Limpiar el puerto de salida del amplificador A-08 y reemplazar el conector corroído.\n6. Sellar con funda termoencogible y cinta autofusionante; dejar el lazo de goteo para que el agua no entre al conector.\n7. Medir de nuevo en el mismo tap y confirmar que no hay fuga de señal.\n8. Seguimiento: sin nuevos reportes de los usuarios del tramo en las 48 horas siguientes, ni siquiera con lluvia.",
    mediciones: [
        { parametro: "Nivel de RF (tap más lejano)", unidad: "dBmV", antes: "2", despues: "11", observacion: "Dentro de lo esperado" },
        { parametro: "MER", unidad: "dB", antes: "22", despues: "37", observacion: "Sin pixelación" },
        { parametro: "Pérdida de retorno", unidad: "dB", antes: "9", despues: "20", observacion: "Línea readaptada" },
        { parametro: "Fuga de señal en el conector", unidad: "Sí / No", antes: "Sí", despues: "No", observacion: "Fuga eliminada" }
    ],
    recursos: [
        { tipo: "Conector", nombre: "Conector .500 nuevo con sello", cantidad: 3, unidad: "und", origen: "Inventario HFC", observacion: "Cantidad estimada" },
        { tipo: "Material", nombre: "Unión recta .500", cantidad: 1, unidad: "und", origen: "Inventario HFC", observacion: "Cantidad estimada" },
        { tipo: "Material", nombre: "Funda termoencogible", cantidad: 3, unidad: "und", origen: "Inventario HFC", observacion: "Cantidad estimada" },
        { tipo: "Material", nombre: "Cinta autofusionante", cantidad: 1, unidad: "rollo", origen: "Inventario HFC", observacion: "Cantidad estimada" }
    ],
    epiAntecedentes: "Troncal instalada hace 9 años con cable coaxial .500 (conductor central de cobre). No hay registro de mantenimiento preventivo en los últimos dos años. El tramo pasa por una zona con árboles y el último reporte de poda fue hace más de un año.",
    epiHallazgos: "- Conector de salida del amplificador flojo, sin funda de sello, con humedad dentro y óxido verdoso (corrosión) en el conductor central y la rosca.\n- Corte de unos 5 cm en el forro del cable, a 60 m del amplificador A-08, por roce con una rama; la malla estaba expuesta y mojada.\n- Mediciones iniciales en el tap más lejano del tramo: nivel de RF 2 dBmV, MER 22 dB, pérdida de retorno 9 dB y fuga de señal detectada en el conector dañado.\n- Se verificó la presencia de voltaje AC en la línea antes de intervenir.",
    epiTratamiento: "Mantenimiento correctivo interno, de alcance menor (se resolvió en sitio el mismo día con materiales de bodega): se cortó y retiró el tramo de cable dañado, se instalaron conectores .500 nuevos con sello y una unión recta, se limpió el puerto del amplificador, se selló todo con funda termoencogible y cinta autofusionante, y se dejó el lazo de goteo para que el agua no entre al conector.",
    epiEvolucion: "Mediciones finales en el mismo tap: nivel de RF 11 dBmV, MER 37 dB, pérdida de retorno 20 dB y sin fuga de señal. Durante las 48 horas siguientes no hubo nuevos reportes de los usuarios del tramo, ni siquiera con lluvia.",
    epiEgreso: "Red operativa y estable, con valores medidos dentro de lo esperado. Orden de trabajo liberada por el Jefe de Operaciones.",
    epiRecomendaciones: "- Inspección preventiva cada 6 meses de conectores, sellos y forro en todos los amplificadores y taps del nodo N-07.\n- Poda de ramas que rozan el cable.\n- Revisar el ajuste (torque) de los conectores y registrar las mediciones en cada visita.\n- Programar una orden de trabajo preventiva para el nodo en los próximos 30 días.",
    jefeMantId: "PE-002",
    jefeAreaId: "PE-003",
    fechaTermino: "2026-10-20"
};

// ---------------------------------------------------------------------------
// CASO 2 - PREVENTIVO: Red Coaxial Externa (OT-2026-0088)
// Hoja "OT Preventivo" del Excel
// ---------------------------------------------------------------------------
const CASO_PREVENTIVO = {
    numOrden: "OT-2026-0088",
    folio: "SOL-2026-0146",
    tipoMantenimiento: "Preventivo",
    tipoArea: "Aérea",
    medioTransmision: "Cable Coaxial",
    clasificacion: "Menor",
    fechaEmision: "2026-10-23",
    fechaProgramacion: "2026-11-12",
    estado: "Abierta",
    tecnicoId: "PE-001",
    clienteId: "CL-001",
    tipoRed: "Red Coaxial Externa (.500-.750)",
    medioEspecificacion: "Cable coaxial .500 (cobre), tramo aéreo de unos 420 m",
    lugarTrabajo: "Nodo N-07: todos los amplificadores y taps del nodo (incluye tramo A-07 – A-08)",
    elementoIntervenido: "Conectores, sellos, forro y malla del cable; amplificadores y taps del nodo N-07",
    proveedorExterno: "No aplica (Personal Interno NOVA NET)",
    motivo: "Mantenimiento preventivo programado del nodo N-07, recomendado en la epicrisis de la OT-2026-0087. No hay registro de mantenimiento preventivo en los últimos dos años y la causa raíz de la falla del 17 al 19 de octubre de 2026 fue la falta de revisión de sellos y forro. No hay falla reportada.",
    diagnostico: "Sin falla reportada; orden programada dentro de los 30 días posteriores a la intervención correctiva del 20/10/2026. Los hallazgos se registran durante la visita: estado de conectores, sellos, forro y malla, humedad o corrosión en los puertos, y ramas que rozan el cable.",
    actividades: "1. Verificar la presencia de voltaje AC en la línea antes de intervenir.\n2. Inspeccionar conectores, sellos y forro de todos los amplificadores y taps del nodo N-07.\n3. Revisar el ajuste (torque) de los conectores y reemplazar sellos o fundas deteriorados.\n4. Revisar los puertos de los amplificadores en busca de humedad o corrosión.\n5. Podar las ramas que rozan el cable en el tramo aéreo.\n6. Medir nivel de RF, MER y pérdida de retorno en los taps y compararlos con los valores de referencia.\n7. Registrar mediciones y hallazgos de la visita; programar la siguiente inspección (cada 6 meses).",
    mediciones: [
        { parametro: "Nivel de RF (tap más lejano)", unidad: "dBmV", antes: "", despues: "", observacion: "Ref.: 11 dBmV" },
        { parametro: "MER", unidad: "dB", antes: "", despues: "", observacion: "Ref.: 37 dB" },
        { parametro: "Pérdida de retorno", unidad: "dB", antes: "", despues: "", observacion: "Ref.: 20 dB" },
        { parametro: "Fuga de señal", unidad: "Sí / No", antes: "", despues: "", observacion: "Ref.: sin fuga" }
    ],
    recursos: [
        { tipo: "Conector", nombre: "Conector .500 con sello (reserva)", cantidad: 4, unidad: "und", origen: "Inventario HFC", observacion: "Cantidad estimada" },
        { tipo: "Material", nombre: "Funda termoencogible", cantidad: 6, unidad: "und", origen: "Inventario HFC", observacion: "Cantidad estimada" },
        { tipo: "Material", nombre: "Cinta autofusionante", cantidad: 1, unidad: "rollo", origen: "Inventario HFC", observacion: "Cantidad estimada" },
        { tipo: "Cable Coaxial", nombre: "Cable coaxial .500 (reserva)", cantidad: 5, unidad: "m", origen: "Inventario HFC", observacion: "Cantidad estimada" }
    ],
    // La epicrisis se completa después de ejecutar la visita
    epiAntecedentes: "", epiHallazgos: "", epiTratamiento: "", epiEvolucion: "", epiEgreso: "", epiRecomendaciones: "",
    jefeMantId: "PE-002",
    jefeAreaId: "PE-003",
    fechaTermino: ""
};

// ---------------------------------------------------------------------------
// PLANTILLA EN BLANCO (sin ninguna información)
// ---------------------------------------------------------------------------
const CASO_VACIO = {
    numOrden: "", folio: "", tipoMantenimiento: "", tipoArea: "", medioTransmision: "", clasificacion: "",
    fechaEmision: "", fechaProgramacion: "", estado: "Abierta",
    tecnicoId: "", clienteId: "", tipoRed: "",
    medioEspecificacion: "", lugarTrabajo: "", elementoIntervenido: "", proveedorExterno: "",
    motivo: "", diagnostico: "", actividades: "",
    mediciones: [], recursos: [],
    epiAntecedentes: "", epiHallazgos: "", epiTratamiento: "", epiEvolucion: "", epiEgreso: "", epiRecomendaciones: "",
    jefeMantId: "", jefeAreaId: "", fechaTermino: ""
};
