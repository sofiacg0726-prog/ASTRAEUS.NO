/* ASTRAEUS | Catálogos de referencia
   Transcripción de las tablas proporcionadas por la autora del proyecto.
   Version de trabajo: las tablas requieren pruebas por producto/campaña
   antes de considerarse equivalentes al 100% a Excel.
*/
(function (root) {
  'use strict';
  const PRODUCTOS = Object.freeze({
    1: { nombre: 'Adquisición tradicional', apertura: 0.0125, prepago: 0.03, ingresoMinimo: 10000, montoMinimo: 250000, montoMaximo: 15000000, viviendaMinima: 400000, pti: 0.55, aforo: 'tabla' },
    2: { nombre: 'Adquisición con Apoyo Infonavit', apertura: 0.0075, prepago: 0, ingresoMinimo: 7500, montoMinimo: 250000, montoMaximo: 15000000, viviendaMinima: 400000, pti: 0.65, aforo: 'tabla' },
    3: { nombre: 'Cofinavit', apertura: 0.0075, prepago: 0, ingresoMinimo: 7500, montoMinimo: 120000, montoMaximo: 15000000, viviendaMinima: 300000, pti: 0.65, aforo: 0.95 },
    4: { nombre: 'Pasivo hipotecario', apertura: 0, prepago: 0.03, ingresoMinimo: 10000, montoMinimo: 125000, montoMaximo: 15000000, viviendaMinima: 0, pti: 0.65, aforo: 'tabla' },
    5: { nombre: 'Pasivo con Apoyo Infonavit', apertura: 0, prepago: 0.03, ingresoMinimo: 10000, montoMinimo: 125000, montoMaximo: 15000000, viviendaMinima: 0, pti: 0.65, aforo: 'tabla' },
    6: { nombre: 'Liquidez vivienda', apertura: 0.0125, prepago: 0.03, ingresoMinimo: 20000, montoMinimo: 150000, montoMaximo: 2500000, viviendaMinima: 1000000, pti: 0.55, aforo: 'tabla' },
    7: { nombre: 'CI+CB Cofinavit', apertura: 0.0125, prepago: 0, ingresoMinimo: 10000, montoMinimo: 250000, montoMaximo: 15000000, viviendaMinima: 300000, pti: 0.55, aforo: 0.90 },
  });
  // Campañas: 1 Ninguna, 2 WM, 3 Premium, 4 Vinculación, 5 Hipoteca Green.
  const PAGOS_OPORTUNOS = Object.freeze({
    1: { tasa: .13, ppm: { 7:18.38, 10:15.05, 15:12.79, 20:11.86 } },
    2: { tasa: .095, ppm: { 7:16.42, 10:13.02, 15:10.53, 20:9.42 } },
    3: { tasa: .0975, ppm: { 7:16.55, 10:13.16, 15:10.70, 20:9.59 } },
    4: { tasa: .1075, ppm: { 7:17.10, 10:13.73, 15:11.32, 20:10.27 } },
    5: { tasa: .10, ppm: { 7:16.70, 10:13.30, 15:10.84, 20:9.75 } },
  });
  // Campaña 1: la tasa, PPM e incremento dependen del índice de riesgo.
  const VALORA_SIN_CAMPANA = Object.freeze({
    1: { tasa: .12, incrementoAnual: .020, ppm: {10:13.63, 15:11.32, 20:10.36} },
    2: { tasa: .125, incrementoAnual: .017, ppm: {10:14.06, 15:11.78, 20:10.84} },
    3: { tasa: .13, incrementoAnual: .014, ppm: {10:14.50, 15:12.22, 20:11.28} },
    4: { tasa: .135, incrementoAnual: .011, ppm: {10:14.90, 15:12.63, 20:11.70} },
    5: { tasa: .14, incrementoAnual: .007, ppm: {10:15.30, 15:13.04, 20:12.13} },
  });
  const VALORA_CAMPANAS = Object.freeze({
    2: { tasa: .095, incrementoAnual: .026, ppm: {10:11.98, 15:9.50, 20:8.38} },
    3: { tasa: .0975, incrementoAnual: .026, ppm: {10:12.12, 15:9.64, 20:8.54} },
    4: { tasa: .1075, incrementoAnual: .024, ppm: {10:12.75, 15:10.33, 20:9.30} },
    5: { tasa: .10, incrementoAnual: .026, ppm: {10:12.24, 15:9.78, 20:8.70} },
  });
  // Tramo 1 hasta $10 M, extraído de la tabla maestra.
  const AFORO_TRAMO1 = Object.freeze({
    'empleado-con-apoyo': [0.95,0.95,0.95,0.85,0.85],
    'empleado-sin-apoyo': [0.95,0.95,0.90,0.85,0.85],
    'independiente': [0.90,0.90,0.85,0.75,0.75],
    'independiente-pasivo': [0.95,0.95,0.95,0.85,0.85],
  });
  // VLOOKUP aproximado de Excel: las claves son límites INFERIORES, no superiores.
  const AVALUOS = Object.freeze([
    [0,1900], [500001,3300], [1000001,4500], [1500001,5750],
    [2000001,7000], [2500001,8200], [3000001,11000], [4000001,13900],
    [5000001,15500], [6000001,19000], [7000001,20000], [8000001,20100],
    [9000001,22500], [10000001,27000], [11000001,28000], [12000001,30000],
    [13000001,32000], [14000001,34000], [15000001,37000], [20000001,39000],
    [25000001,40000], [30000001,40000]
  ]);
  function condicionesTasa({ esquema, campana, plazoAnios, indiceRiesgo = 1 }) {
    if (esquema === 'pagos-oportunos') {
      const entrada = PAGOS_OPORTUNOS[campana];
      if (!entrada || entrada.ppm[plazoAnios] === undefined) throw new Error('Campaña o plazo no válido para Pagos Oportunos');
      return { tasaInicial: entrada.tasa, ppm: entrada.ppm[plazoAnios], incrementoAnual: 0 };
    }
    if (esquema === 'valora') {
      const entrada = campana === 1 ? VALORA_SIN_CAMPANA[indiceRiesgo] : VALORA_CAMPANAS[campana];
      if (!entrada || entrada.ppm[plazoAnios] === undefined) throw new Error('Campaña, IR o plazo no válido para Valora');
      return { tasaInicial: entrada.tasa, ppm: entrada.ppm[plazoAnios], incrementoAnual: entrada.incrementoAnual };
    }
    throw new Error('Esquema desconocido');
  }
  function aforoDeReferencia({ productoId, tipoCliente, indiceRiesgo, montoReferencia }) {
    const prod = PRODUCTOS[productoId];
    if (!prod) throw new Error('Producto no reconocido');
    if (typeof prod.aforo === 'number') return prod.aforo;
    if (!Number.isInteger(indiceRiesgo) || indiceRiesgo < 1 || indiceRiesgo > 5) throw new Error('Índice de riesgo inválido');
    if (!Number.isFinite(montoReferencia) || montoReferencia < 0) throw new Error('Monto de referencia inválido');
    if (montoReferencia > 10000000) {
      throw new Error('Aforo de más de $10 millones pendiente de conciliar con la referencia defectuosa I21/AH48:AJ52. No se aproxima.');
    }
    // Solo los productos 2 y 5 se consideran "con apoyo" en la regla aportada.
    const categoria = (productoId === 2 || productoId === 5) && tipoCliente === 'empleado'
      ? 'empleado-con-apoyo'
      : tipoCliente === 'empleado' ? 'empleado-sin-apoyo'
      : (productoId === 4 || productoId === 5) ? 'independiente-pasivo' : 'independiente';
    return AFORO_TRAMO1[categoria][indiceRiesgo - 1];
  }
  function avaluoPorValor(valor) {
    if (!Number.isFinite(valor) || valor < 0) throw new Error('Valor de la vivienda inválido');
    // Reproduce la búsqueda aproximada de VLOOKUP.
    // $500,000.50 todavía cae en la banda cuyo límite inferior es $0.
    let importe = AVALUOS[0][1];
    for (const [desde, cuota] of AVALUOS) {
      if (valor < desde) break;
      importe = cuota;
    }
    return importe;
  }
  const api = { PRODUCTOS, PAGOS_OPORTUNOS, VALORA_SIN_CAMPANA, VALORA_CAMPANAS,
    AFORO_TRAMO1, AVALUOS, condicionesTasa, aforoDeReferencia, avaluoPorValor };
  root.AstraeusCatalogos = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
