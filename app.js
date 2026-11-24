/* ASTRAEUS — interfaz del simulador principal
   Conecta el HTML del proyecto con los catálogos y el motor financiero.
   Fórmulas de referencia: Excel .xlsm suministrado por la autora.
   NO afirmar equivalencia total: caso guardado verificado, demás escenarios por probar.
*/
(function () {
  'use strict';
  document.addEventListener('DOMContentLoaded', iniciar);

  function iniciar() {
    const C = window.AstraeusCatalogos;
    const F = window.AstraeusFinanzas;
    if (!C || !F) {
      alert('ASTRAEUS: no se cargaron js/catalogos.js y js/motor-financiero.js.');
      return;
    }
    const $ = id => document.getElementById(id);
    const CAMPANAS = {ninguna:1,wm:2,premium:3,vinculacion:4,green:5};
    const NOMBRES = {1:'Ninguna',2:'WM',3:'Premium',4:'Vinculación',5:'Hipoteca Green'};
    const CODIGOS = {1:'ninguna',2:'wm',3:'premium',4:'vinculacion',5:'green'};
    const PISOS_VALORA = {1:.12,2:.095,3:.0975,4:.1075,5:.10};
    let escenarioActivo = null;

    function num(id, obligatorio=false) {
      const el = $(id);
      if (!el) throw new Error(`Falta el campo #${id} en el HTML`);
      const raw = el.value.trim();
      if (raw === '') {
        if (obligatorio) throw new Error(`Ingresa un valor para «${el.labels?.[0]?.textContent?.trim() || id}».`);
        return undefined;
      }
      const n = Number(raw);
      if (!Number.isFinite(n) || n < 0) throw new Error(`El campo «${id}» debe contener un importe válido y no negativo.`);
      if (obligatorio && n <= 0) throw new Error(`El campo «${id}» debe ser mayor que cero.`);
      return n;
    }
    function moneda(n) { return F.pesos(n); }
    function pct(decimal, digitos=2) { return (decimal*100).toFixed(digitos)+'%'; }
    function texto(id, valor) { const el=$(id); if(el) el.textContent=valor; }
    function errorTexto(e) { return e instanceof Error ? e.message : String(e); }

    function mostrarMensajes(errores=[], advertencias=[]) {
      const panel = $('mensajesSimulacion');
      if (!panel) return;
      panel.replaceChildren();
      panel.hidden = errores.length===0 && advertencias.length===0;
      panel.classList.toggle('error', errores.length>0);
      if (panel.hidden) return;
      const encabezado=document.createElement('strong');
      encabezado.textContent=errores.length ? 'No se pudo calcular' : 'Observaciones de la simulación';
      panel.append(encabezado);
      const lista=document.createElement('ul');
      [...errores,...advertencias].forEach(msg=>{
        const li=document.createElement('li');li.textContent=msg;lista.appendChild(li);
      });
      panel.append(lista);
    }

    function estadoFormulario() {
      const modalidad=$('tipoCalculo').value;
      $('campoValor').hidden = modalidad !== 'valor';
      $('campoIngreso').hidden = modalidad !== 'ingresos';
      $('campoCredito').hidden = modalidad !== 'credito';
      const idProd=Number($('producto').value);
      $('campoCotizacion').hidden = ![2,5].includes(idProd);
      $('campoSubcuenta').hidden = ![3,7].includes(idProd);
      // Cuando Apoyo Infonavit exige el salario, mostramos automáticamente el campo.
      if ([2,5].includes(idProd) && $('opcionesSeguroApoyo')) {
        $('opcionesSeguroApoyo').open = true;
      }
      llenarCatalogoTasas();
    }
    function condiciones(campana) {
      const esquema=$('esquema').value==='pagos'?'pagos-oportunos':'valora';
      const plazoAnios=Number($('plazo').value);
      const indiceRiesgo=Number($('riesgo').value);
      const tasas=C.condicionesTasa({esquema,campana,plazoAnios,indiceRiesgo});
      return {esquema,plazoAnios,indiceRiesgo,...tasas,
        pisoValora:esquema==='valora'?PISOS_VALORA[campana]:undefined};
    }
    function llenarCatalogoTasas() {
      const select=$('catalogoTasa');
      if(!select) return;
      const elegida=CAMPANAS[$('campania').value] ?? 5;
      select.replaceChildren();
      let cantidad=0;
      for(const campana of [1,2,3,4,5]) {
        try {
          const c=condiciones(campana);
          const op=document.createElement('option');
          op.value=String(campana);
          op.textContent=`${NOMBRES[campana]} · ${pct(c.tasaInicial)} · PPM $${c.ppm.toFixed(2)}`;
          select.appendChild(op); cantidad++;
        } catch (_error) { /* Plazo no ofrecido: no se muestra como opción. */ }
      }
      if (cantidad && [...select.options].some(o=>Number(o.value)===elegida)) select.value=String(elegida);
      if (!cantidad) { const op=document.createElement('option');op.textContent='No disponible para este plazo';op.value='';select.appendChild(op); }
    }
    function fechaHoyLocal() {
      const d=new Date();
      return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    }

    // Parte 1: cálculo desde los tres modos del panel Excel.
    function crearEscenarioNuevo() {
      const productoId=Number($('producto').value);
      if (!C.PRODUCTOS[productoId]) throw new Error('Selecciona un producto válido.');
      const producto=C.PRODUCTOS[productoId];
      const modalidad=$('tipoCalculo').value;
      const campana=CAMPANAS[$('campania').value];
      const c=condiciones(campana);
      const tipoCliente=$('actividad').value;
      if ([2,5].includes(productoId) && tipoCliente!=='empleado') {
        throw new Error('Los productos con Apoyo Infonavit están configurados para empleados; revisa la combinación elegida.');
      }
      const tipoComision=Number($('tipoComision').value);
      const valorIngresado=num('valorVivienda',modalidad==='valor');
      const ingresoIngresado=num('ingreso',modalidad==='ingresos');
      const creditoDeseado=num('montoCredito',modalidad==='credito');
      const sueldoCotizacionMensual=num('sueldoCotizacion',[2,5].includes(productoId));
      const valorAsegurable=num('valorAsegurable');
      const subcuenta=num('subcuenta')??0;
      const fechaReferencia=$('fechaReferencia').value;
      const segurosExternos=$('segurosExternos').value==='si';
      if (!fechaReferencia) throw new Error('Selecciona la fecha de referencia.');
      if (modalidad==='credito' && creditoDeseado>10000000) throw new Error('Tramo superior a $10 millones aún no conciliado con las referencias del Excel.');
      const montoParaAforo=modalidad==='valor' ? valorIngresado : modalidad==='credito' ? creditoDeseado : (ingresoIngresado*producto.pti)/(c.ppm/1000);
      const aforo=C.aforoDeReferencia({productoId,tipoCliente,indiceRiesgo:c.indiceRiesgo,montoReferencia:montoParaAforo});
      const resultado=C.PRODUCTOS[productoId];
      const base=F.calcularPorModalidad({modalidad,valorVivienda:valorIngresado,ingresoMensual:ingresoIngresado,
        creditoDeseado,aforo,pti:resultado.pti,ppm:c.ppm,productoId});
      if (base.creditoTotal>10000000) throw new Error('El modelo de aforo para montos superiores a $10 millones continúa en validación y no se estimará automáticamente.');
      const distribucion=F.distribuirCofinavit({creditoTotal:base.creditoTotal,productoId});
      return {
        productoId,productoNombre:producto.nombre,tipoCliente,modalidad,aforo,pti:producto.pti,
        creditoTotal:distribucion.creditoTotal,creditoBanco:distribucion.creditoBanco,
        creditoInfonavit:distribucion.creditoInfonavit,
        valorVivienda:base.valorViviendaEstimado, ingresoMensual:ingresoIngresado,
        ingresoNecesario:base.ingresoNecesario,
        campana,tipoComision,segurosExternos,valorAsegurable,subcuenta,fechaReferencia,
        sueldoCotizacionMensual, advertencias:base.advertencias, ...c,
        prepagos:leerPrepagos(),
      };
    }
    function leerPrepagos() {
      const mes=num('mesPrepago');
      const importe=num('montoPrepago');
      if (importe===undefined || importe===0) {
        if (mes!==undefined) throw new Error('Si señalas un mes de prepago, indica su importe.');
        return {};
      }
      if (!Number.isInteger(mes) || mes<1 || mes>240) throw new Error('Indica un mes de prepago entre 1 y 240.');
      return {[mes]:importe};
    }

    // Parte 2: cálculo financiero, gastos, amortización y CAT.
    function ejecutarEscenario(p) {
      const gastos=F.calcularGastos({productoId:p.productoId,campana:p.campana,tipoComision:p.tipoComision,
        valorVivienda:p.valorVivienda,creditoParaApertura:p.creditoBanco,creditoInfonavit:p.creditoInfonavit});
      const resumen=F.resumenOperacion({valorVivienda:p.valorVivienda,recursosBanco:p.creditoBanco,
        creditoInfonavit:p.creditoInfonavit,subcuenta:p.subcuenta,gastos});
      const cfg={...p,porcentajePenalizacion:C.PRODUCTOS[p.productoId].prepago,
        comisionApertura:gastos.apertura,avaluo:gastos.avaluo,modalidadPrepago:'plazo'};
      const tabla=F.simularTabla(cfg);
      let cat=null, errorCat=null;
      try { cat=F.calcularCAT(cfg); } catch(e) {errorCat=errorTexto(e);}
      const notas=[...(p.advertencias||[])];
      if ([3,7].includes(p.productoId)) notas.push('Cofinavit: la división del crédito utiliza el supuesto del 30% y el tope guardado en el Excel de referencia.');
      if ([4,5,6].includes(p.productoId)) notas.push('Pasivos/Liquidez: aún falta verificar contra Excel las condiciones especiales de financiamiento de gastos y aplicación de recursos.');
      if (p.esquema==='valora') notas.push('Valora: los pisos y reglas especiales deben cotejarse con nuevas corridas de Excel.');
      if (p.creditoBanco<0.01) throw new Error('El crédito bancario calculado no puede ser cero.');
      if (tabla.ultimoSaldo>0.01) notas.push('El crédito no se liquidó en los 240 meses modelados; revisar tasa, PPM o condiciones especiales.');
      if (resumen.enganche<0) notas.push('Los recursos estimados superan el valor de vivienda: revisar subcuenta y distribución del crédito.');
      if (errorCat) notas.push('CAT no calculable en este escenario: '+errorCat);
      return {p,gastos,resumen,tabla,cat,notas};
    }

    function render(result) {
      const {p,gastos,resumen,tabla,cat}=result;
      const primera=tabla.filas[0];
      texto('creditoResultado',moneda(p.creditoBanco));
      texto('pagoResultado',moneda(primera.pagoTotal));
      texto('tasaResultado',pct(p.tasaInicial));
      texto('ppmResultado',moneda(p.ppm));
      texto('aforoResultado',pct(p.aforo,0));
      texto('catResultado',cat?pct(cat.catEfectivoAnual,3):'No disponible');
      texto('totalPagadoResultado',moneda(tabla.totalPagado));
      texto('plazoResultado',`${p.plazoAnios} años / ${Math.floor(tabla.pagos/12)} años y ${tabla.pagos%12} meses`);
      texto('infonavitResultado',moneda(p.creditoInfonavit));
      texto('engancheResultado',moneda(resumen.enganche));
      texto('inversionResultado',moneda(resumen.inversionCliente));
      texto('avaluo',moneda(gastos.avaluo));
      texto('impuestos',moneda(gastos.impuestos));
      texto('notario',moneda(gastos.notario));
      texto('comision',moneda(gastos.apertura));
      texto('comisionInfonavit',moneda(gastos.comisionInfonavit));
      texto('gastosTotal',moneda(gastos.total));
      $('catalogoTasa').value=String(p.campana);
      const cuerpo=$('tablaBody');cuerpo.replaceChildren();
      const fragment=document.createDocumentFragment();
      for(const r of tabla.filas) {
        const tr=document.createElement('tr');
        const valores=[r.mes,r.periodo,moneda(r.saldoInicial),pct(r.tasaAplicada),moneda(r.interes),
          moneda(r.erogacion),moneda(r.amortizacionCapital),moneda(r.prepago),moneda(r.aportacion),
          moneda(r.seguroVida+r.seguroDanos),moneda(r.pagoTotal),moneda(r.saldoFinal)];
        valores.forEach(v=>{const td=document.createElement('td');td.textContent=String(v);tr.appendChild(td);});
        fragment.appendChild(tr);
      }
      cuerpo.appendChild(fragment);
      mostrarMensajes([],result.notas);
    }
    $('btnCalcular').addEventListener('click',()=>{
      try {
        const p=crearEscenarioNuevo();
        const r=ejecutarEscenario(p);
        escenarioActivo=p;
        render(r);
      } catch(e) {mostrarMensajes([errorTexto(e)]);}
    });
    $('btnRecalcular').addEventListener('click',()=>{
      if(!escenarioActivo) {mostrarMensajes(['Calcula primero tu crédito para poder recalcularlo.']);return;}
      try {
        const nuevaCampana=Number($('catalogoTasa').value);
        if(!nuevaCampana) throw new Error('Selecciona una campaña/tasa del catálogo.');
        const c=C.condicionesTasa({esquema:escenarioActivo.esquema,campana:nuevaCampana,
          plazoAnios:escenarioActivo.plazoAnios,indiceRiesgo:escenarioActivo.indiceRiesgo});
        // Se mantiene exactamente el crédito del escenario ya obtenido.
        const p={...escenarioActivo,...c,campana:nuevaCampana,
          pisoValora:escenarioActivo.esquema==='valora'?PISOS_VALORA[nuevaCampana]:undefined};
        const r=ejecutarEscenario(p);
        escenarioActivo=p;
        $('campania').value=CODIGOS[nuevaCampana];
        render(r);
      } catch(e) {mostrarMensajes([errorTexto(e)]);}
    });

    $('btnAceptarAviso')?.addEventListener('click',()=>{
      $('avisoModal').style.display='none';
    });
    ['tipoCalculo','producto','esquema','plazo','riesgo','campania'].forEach(id=>{
      $(id).addEventListener('change',estadoFormulario);
    });
    if(!$('fechaReferencia').value) $('fechaReferencia').value=fechaHoyLocal();
    estadoFormulario();
  }
})();
