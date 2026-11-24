/* ASTRAEUS | MOTOR FINANCIERO v1 (compatible con la corrida de referencia)
   Fundamento: amortización sobre saldo insoluto, interés ordinario Actual/360,
   pago por millar contractual, tasas escalonadas y CAT como TIR de flujos mensuales.
   El motor NO redondea los valores intermedios a centavos.
   Importante: modalidad "disminución de pago" tras prepago y pisos de Valora
   requieren una calibración adicional contra el Excel original.
*/
(function (root) {
  'use strict';
  const CAT = root.AstraeusCatalogos || (typeof require === 'function' ? require('./catalogos.js') : null);
  const CONSTANTES = Object.freeze({
    baseIntereses: 360,
    seguroVidaMensual: 0.0006,
    seguroDanosMensual: 0.00031,
    ivaPenalizacion: 0.16,
    aportacionPatronal: 0.05,
    umaDiariaReferencia2026: 117.31,
    diasBimestre: 60.8,
    crecimientoAportacionAnual: 0.03,
    porcentajeInfonavitCofinavit: 0.30,
    topeInfonavitReferencia: 885445.60,
    costoOriginacionInfonavit: 0.03,
  });
  function exigirNumero(nombre, numero, min = 0) {
    if (!Number.isFinite(numero) || numero < min) throw new Error(`${nombre} debe ser numérico y no menor a ${min}`);
    return numero;
  }
  function parsearFechaISO(fecha) {
    if (typeof fecha !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('fechaReferencia debe tener formato AAAA-MM-DD');
    const [y,m,d] = fecha.split('-').map(Number);
    const prueba = new Date(Date.UTC(y,m-1,d));
    if (prueba.getUTCFullYear()!==y || prueba.getUTCMonth()!==m-1 || prueba.getUTCDate()!==d) throw new Error('fechaReferencia no es una fecha real');
    return { y, m, d };
  }
  function primerosDatosPago(fecha) {
    const {y,m,d} = parsearFechaISO(fecha);
    const diaCorte = d >= 3 && d <= 16 ? 3 : 17;
    const sumarMes = d > diaCorte ? 1 : 0;
    const fechaPago = new Date(Date.UTC(y,m-1+sumarMes,diaCorte));
    return { anio: fechaPago.getUTCFullYear(), mes: fechaPago.getUTCMonth()+1, dia: diaCorte };
  }
  function periodoPago(primero, numeroMes) {
    const f = new Date(Date.UTC(primero.anio, primero.mes - 1 + numeroMes - 1, 1));
    const y = f.getUTCFullYear(), m = f.getUTCMonth() + 1;
    const dias = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return { anio:y, mes:m, dias, claveMes: `${y}-${String(m).padStart(2,'0')}` };
  }
  function tasaDelMes({ esquema, tasaInicial, mes, pisoValora }) {
    if (esquema === 'pagos-oportunos') {
      // Cambios de puntos porcentuales EXACTOS según la tabla suministrada.
      const rebaja = mes >= 49 ? .0075 : mes >= 37 ? .006 : mes >= 25 ? .004 : mes >= 13 ? .002 : 0;
      return tasaInicial - rebaja;
    }
    if (esquema === 'valora') {
      if (!Number.isFinite(pisoValora)) throw new Error('Para Valora se requiere pisoValora verificado en el Excel');
      const escalones = mes >= 37 ? Math.floor((mes - 37)/12) + 1 : 0;
      return Math.max(pisoValora, tasaInicial - .0025 * escalones);
    }
    throw new Error('Esquema desconocido');
  }
  function aumentoPagoMes(esquema, incrementoAnual, mes) {
    return esquema === 'valora' && mes >= 13 && (mes - 1) % 12 === 0 ? incrementoAnual : 0;
  }
  function aportacionBimestral({ productoId, mes, saldoInicial, sueldoCotizacionMensual }) {
    if (![2,5].includes(productoId) || mes % 2 !== 0 || saldoInicial <= 0) return 0;
    if (!Number.isFinite(sueldoCotizacionMensual)) throw new Error('Falta sueldoCotizacionMensual para Apoyo Infonavit');
    const topeInicial = CONSTANTES.umaDiariaReferencia2026 * 25 * CONSTANTES.aportacionPatronal * CONSTANTES.diasBimestre;
    const base = Math.min(sueldoCotizacionMensual * CONSTANTES.aportacionPatronal * 2, topeInicial);
    // Se usa como hipótesis de trabajo hasta validar cada transición de año de la col. N.
    const anios = Math.floor((mes - 1) / 12);
    return base * (1 + CONSTANTES.crecimientoAportacionAnual) ** anios;
  }
  function calcularGastos({ productoId, campana, tipoComision, valorVivienda, creditoParaApertura,
      creditoInfonavit = 0, impuestosTasa = 0.03, notarioTasa = 0.02,
      avaluoOverride, otrosGastos = 0, gastosFinanciados = 0 }) {
    if (!CAT) throw new Error('Se necesita cargar catalogos.js antes del motor');
    const producto = CAT.PRODUCTOS[productoId];
    if (!producto) throw new Error('Producto no válido');
    [ ['valorVivienda',valorVivienda], ['creditoParaApertura',creditoParaApertura], ['creditoInfonavit',creditoInfonavit] ]
      .forEach(([n,v])=>exigirNumero(n,v));
    const avaluo = avaluoOverride !== undefined ? exigirNumero('avaluoOverride', avaluoOverride)
      : ([2,3].includes(campana) || [4,5].includes(productoId) ? 0 : CAT.avaluoPorValor(valorVivienda));
    const apertura = (tipoComision === 1 && campana === 1) ? creditoParaApertura*producto.apertura : 0;
    const comisionInfonavit = [3,7].includes(productoId) ? creditoInfonavit * CONSTANTES.costoOriginacionInfonavit : 0;
    const impuestos = valorVivienda * impuestosTasa;
    const notario = valorVivienda * notarioTasa;
    const total = impuestos + notario + avaluo + apertura + comisionInfonavit + otrosGastos - gastosFinanciados;
    return { impuestos, notario, avaluo, apertura, comisionInfonavit, otrosGastos, gastosFinanciados, total };
  }
  function resumenOperacion({ valorVivienda, recursosBanco, creditoInfonavit=0, subcuenta=0, gastos }) {
    const enganche = valorVivienda - recursosBanco - creditoInfonavit - subcuenta;
    const inversionCliente = enganche + gastos.total;
    return { enganche, inversionCliente };
  }
  function calcularPorModalidad({ modalidad, valorVivienda, ingresoMensual, creditoDeseado,
      aforo, pti, ppm, productoId }) {
    // Cálculo financiero por 3 modalidades del panel Y41.
    // No sustituye las validaciones ni la lógica condicionada de las celdas W82/W83.
    exigirNumero('aforo', aforo, 0.000001);
    if (aforo>1) throw new Error('aforo no puede superar 100%');
    exigirNumero('ppm', ppm, 0.000001);
    if (pti !== undefined) exigirNumero('pti', pti, 0.000001);
    let creditoTotal, valorInferido, ingresoInferido;
    if (modalidad==='valor') {
      exigirNumero('valorVivienda', valorVivienda, 0.01);
      creditoTotal=valorVivienda*aforo;
      valorInferido=valorVivienda;
    } else if (modalidad==='ingresos') {
      exigirNumero('ingresoMensual', ingresoMensual, 0.01);
      if(!Number.isFinite(pti)) throw new Error('Falta PTI');
      creditoTotal=(ingresoMensual*pti)/(ppm/1000);
      valorInferido=creditoTotal/aforo;
    } else if (modalidad==='credito') {
      exigirNumero('creditoDeseado', creditoDeseado, 0.01);
      creditoTotal=creditoDeseado;
      valorInferido=creditoTotal/aforo;
    } else throw new Error('modalidad debe ser valor, ingresos o credito');
    ingresoInferido=Number.isFinite(pti) ? (creditoTotal/1000*ppm)/pti : null;
    const producto=CAT.PRODUCTOS[productoId];
    const advertencias=[];
    if(producto){
      if(creditoTotal<producto.montoMinimo) advertencias.push('Crédito por debajo del mínimo del producto');
      if(creditoTotal>producto.montoMaximo) advertencias.push('Crédito por encima del máximo del producto');
      if(valorInferido<producto.viviendaMinima) advertencias.push('Vivienda por debajo del mínimo del producto');
      if(ingresoMensual!==undefined && ingresoMensual<producto.ingresoMinimo) advertencias.push('Ingreso inferior al mínimo del producto');
    }
    return {creditoTotal, valorViviendaEstimado:valorInferido, ingresoNecesario:ingresoInferido,
      pagoBase:creditoTotal/1000*ppm, advertencias};
  }
  function distribuirCofinavit({creditoTotal, productoId, topeInfonavit=CONSTANTES.topeInfonavitReferencia}) {
    exigirNumero('creditoTotal',creditoTotal);
    const creditoInfonavit=[3,7].includes(productoId)
      ? Math.min(creditoTotal*CONSTANTES.porcentajeInfonavitCofinavit, topeInfonavit) : 0;
    return {creditoTotal,creditoInfonavit,creditoBanco:creditoTotal-creditoInfonavit};
  }
  function simularTabla(parametros) {
    const p = parametros;
    const saldoInicialOriginal = exigirNumero('creditoBanco',p.creditoBanco,0.01);
    const valorVivienda = exigirNumero('valorVivienda',p.valorVivienda,0);
    const ppm = exigirNumero('ppm',p.ppm,0.001);
    const tasaInicial = exigirNumero('tasaInicial',p.tasaInicial,0);
    const incrementoAnual = p.incrementoAnual || 0;
    const primero = primerosDatosPago(p.fechaReferencia);
    const n = p.maxMeses ?? 240;
    if (!Number.isInteger(n) || n < 1 || n > 600) throw new Error('maxMeses inválido');
    if (p.modalidadPrepago === 'pago') throw new Error('Disminución del pago pendiente de replicar la fórmula AQ del Excel (no se usa aproximación).');
    const amortizacion = [];
    let saldo = saldoInicialOriginal;
    let pagoAnterior = null;
    for (let mes = 1; mes <= n; mes++) {
      if (saldo <= 0) break;
      const periodo = periodoPago(primero,mes);
      const tasaAplicada = tasaDelMes({ esquema:p.esquema, tasaInicial, mes, pisoValora:p.pisoValora });
      const interes = tasaAplicada / CONSTANTES.baseIntereses * periodo.dias * saldo;
      const aumento = aumentoPagoMes(p.esquema, incrementoAnual, mes);
      let erogacion;
      if (mes === 1) erogacion = saldoInicialOriginal / 1000 * ppm;
      else if (saldo + interes - pagoAnterior <= 0) erogacion = saldo + interes;
      else erogacion = pagoAnterior * (1 + aumento);
      const amortizacionCapital = erogacion - interes;
      const prepago = p.prepagos?.[mes] || 0;
      exigirNumero(`prepago del mes ${mes}`, prepago);
      const penalizacion = p.tipoComision === 2 ? prepago * (p.porcentajePenalizacion || 0) * (1 + CONSTANTES.ivaPenalizacion) : 0;
      const aportacion = p.sinAportaciones ? 0 : aportacionBimestral({
        productoId:p.productoId, mes, saldoInicial:saldo, sueldoCotizacionMensual:p.sueldoCotizacionMensual });
      const seguroVida = p.segurosExternos ? 0 : saldo * (p.tasaSeguroVida ?? CONSTANTES.seguroVidaMensual);
      const valorAsegurable = p.valorAsegurable > 0 ? p.valorAsegurable : valorVivienda;
      const seguroDanos = p.segurosExternos ? 0 : valorAsegurable * (p.tasaSeguroDanos ?? CONSTANTES.seguroDanosMensual);
      const pagoTotal = erogacion + prepago + seguroVida + seguroDanos;
      const saldoFinal = saldo - amortizacionCapital - prepago + penalizacion - aportacion;
      amortizacion.push({ mes, periodo:periodo.claveMes, dias:periodo.dias, saldoInicial:saldo,
        tasaAplicada, erogacion, interes, amortizacionCapital, prepago, penalizacion, aportacion,
        seguroVida, seguroDanos, pagoTotal, saldoFinal });
      pagoAnterior = erogacion;
      saldo = saldoFinal <= 0 ? 0 : saldoFinal;
    }
    return { filas:amortizacion, pagos:amortizacion.length,
      totalPagado:amortizacion.reduce((a,f)=>a+f.pagoTotal,0),
      totalIntereses:amortizacion.reduce((a,f)=>a+f.interes,0),
      ultimoSaldo:saldo, primerPago:primero };
  }
  function calcularTIRMensual(flujoInicialNeto, pagos) {
    exigirNumero('flujoInicialNeto',flujoInicialNeto,0.000001);
    if (!Array.isArray(pagos) || !pagos.length) throw new Error('Se necesita una secuencia de pagos');
    function vpn(t) {
      let suma = -flujoInicialNeto;
      for (let k=0; k<pagos.length; k++) suma += pagos[k] / (1+t)**(k+1);
      return suma;
    }
    if (vpn(0) < 0) throw new Error('La corrida no tiene una TIR positiva con estos flujos');
    let inf=0, sup=.05;
    while(vpn(sup)>0 && sup<1024) sup=sup*2+0.05;
    if (vpn(sup)>0) throw new Error('No se pudo acotar la TIR');
    for(let k=0; k<120; k++) {
      const medio=(inf+sup)/2;
      if(vpn(medio)>0) inf=medio; else sup=medio;
    }
    return (inf+sup)/2;
  }
  function calcularCAT(parametros) {
    // El archivo lleva una segunda corrida sin prepagos para los flujos del CAT.
    const sinPrepagos = simularTabla({...parametros, prepagos:{}, sinAportaciones:true, modalidadPrepago:'plazo'});
    const creditoNeto = parametros.creditoBanco - (parametros.comisionApertura ?? 0) - (parametros.avaluo ?? 0);
    if(creditoNeto<=0) throw new Error('Monto neto para CAT debe ser positivo');
    const tirMensual = calcularTIRMensual(creditoNeto,sinPrepagos.filas.map(f=>f.pagoTotal));
    return { tirMensual, catEfectivoAnual:(1+tirMensual)**12-1, flujoInicialNeto:creditoNeto,
      pagosParaCAT:sinPrepagos.pagos };
  }
  function pesos(n) { return new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',minimumFractionDigits:2,maximumFractionDigits:2}).format(n); }
  const api = { CONSTANTES, primerosDatosPago, periodoPago, tasaDelMes, calcularGastos, calcularPorModalidad, distribuirCofinavit,
    resumenOperacion, simularTabla, calcularTIRMensual, calcularCAT, pesos };
  root.AstraeusFinanzas=api;
  if(typeof module !== 'undefined' && module.exports) module.exports=api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
