// ============================================================
//  AGA — AYARZA GARRE AUTOS  |  v2.1
//  Caja hub · Reportes · Personas
// ============================================================
const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── Entry / Auth ─────────────────────────────────────────────
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.userInfo = getUserInfo_();
  return t.evaluate()
    .setTitle('AGA — Ayarza Garre Autos')
    .addMetaTag('viewport','width=device-width,initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }

function getUserInfo_() {
  try {
    const email = Session.getActiveUser().getEmail();
    const sh = SS.getSheetByName('CONFIG');
    if (sh) {
      for (const r of sh.getDataRange().getValues())
        if (String(r[0]).toLowerCase() === email.toLowerCase())
          return { email, role: r[1], nombre: r[2] || email };
      const owner = DriveApp.getFileById(SS.getId()).getOwner().getEmail();
      if (email === owner) return { email, role: 'admin', nombre: 'Admin' };
    }
    return { email, role: 'vendedor', nombre: email };
  } catch (e) { return { email: 'demo@aga.com', role: 'admin', nombre: 'Demo' }; }
}
function getUserInfo() { return getUserInfo_(); }

// ── Dashboard ─────────────────────────────────────────────────
function getDashboardData() {
  try {
    const now = new Date(), fom = new Date(now.getFullYear(), now.getMonth(), 1);
    let enStock = 0, reservados = 0, consignas = 0, consignasMartin = 0;
    const vSh = SS.getSheetByName('VEHICULOS');
    if (vSh && vSh.getLastRow() > 1)
      vSh.getRange(2,1,vSh.getLastRow()-1,16).getValues().forEach(r => {
        if (!r[0] || r[6]==='Eliminado') return;
        if (r[6]==='Stock') enStock++;
        if (r[6]==='Reservado') reservados++;
        if (r[5]==='Consigna' && r[6]!=='Vendido'){
          consignas++;
          // ex_dueno_nombre='Martin' indica que es auto de Martín
          if(String(r[11]||'').toLowerCase().includes('martin')) consignasMartin++;
        }
      });
    let ventasMes = 0, gananciaMes = 0; const ventasRecientes = [];
    const vtSh = SS.getSheetByName('VENTAS');
    if (vtSh && vtSh.getLastRow() > 1) {
      const rows = vtSh.getRange(2,1,vtSh.getLastRow()-1,15).getValues();
      rows.forEach(r => { if (!r[0]) return; if (new Date(r[4])>=fom){ventasMes++;gananciaMes+=Number(r[13])||0;} });
      rows.filter(r=>r[0]).sort((a,b)=>new Date(b[4])-new Date(a[4])).slice(0,5).forEach(r =>
        ventasRecientes.push({vehiculo:r[3],comprador:r[5],precio:r[7],ganancia:r[13],fecha:_fmt(r[4],'dd/MM/yy')}));
    }
    let saldoCaja = 0;
    const cSh = SS.getSheetByName('CAJA');
    if (cSh && cSh.getLastRow() > 1) saldoCaja = cSh.getRange(cSh.getLastRow(),9).getValue()||0;
    const chequesPorVencer = [];
    const chSh = SS.getSheetByName('CHEQUES');
    if (chSh && chSh.getLastRow() > 1) {
      const en15 = new Date(now.getTime()+15*864e5);
      const chCols = Math.min(chSh.getLastColumn(),13);
      chSh.getRange(2,1,chSh.getLastRow()-1,chCols).getValues().forEach(r => {
        if (!r[0]) return;
        if ((r[12]||'AGA')!=='AGA') return; // solo alertas de AGA en dashboard
        const vto = new Date(r[7]); // col 8 = índice 7 = FECHA_VTO (fix: era r[8])
        if (r[9]==='En cartera'&&vto<=en15&&vto>=now)
          chequesPorVencer.push({numero:r[1],banco:r[3],librador:r[4],importe:r[6],
            fecha:_fmt(vto,'dd/MM/yy'),dias:Math.ceil((vto-now)/864e5)});
      });
      chequesPorVencer.sort((a,b)=>a.dias-b.dias);
    }
    let tcBlue = 0, tcOficial = 0;
    SS.getSheetByName('CONFIG')?.getDataRange().getValues().forEach(r=>{
      if(r[0]==='tc_blue') tcBlue=r[1]; if(r[0]==='tc_oficial') tcOficial=r[1];
    });
    return {ok:true,enStock,reservados,consignas,consignasMartin,ventasMes,gananciaMes,saldoCaja,
      chequesPorVencer,tcBlue,tcOficial,ventasRecientes,
      mes:Utilities.formatDate(now,'America/Argentina/Buenos_Aires','MMMM yyyy')};
  } catch(e){ return {ok:false,error:e.message}; }
}

// ── Config / Listas ──────────────────────────────────────────
function getConfigListas() {
  const sh = SS.getSheetByName('CONFIG');
  const l = {categorias_gasto:[],medios_pago:[],tc_blue:0,tc_oficial:0,marcas:[]};
  if (!sh) return l;
  sh.getDataRange().getValues().forEach(r => {
    if (r[0]==='tc_blue')    l.tc_blue   = r[1];
    if (r[0]==='tc_oficial') l.tc_oficial = r[1];
    if (String(r[0]).startsWith('CAT_GASTO_')) l.categorias_gasto.push(r[1]);
    if (String(r[0]).startsWith('MDP_'))       l.medios_pago.push(r[1]);
    if (String(r[0]).startsWith('MARCA_'))     l.marcas.push(r[1]);
  });
  l.marcas.sort();
  return l;
}
function actualizarTC(b,o) {
  const sh = SS.getSheetByName('CONFIG'); if (!sh) return {ok:false};
  sh.getDataRange().getValues().forEach((r,i)=>{
    if(r[0]==='tc_blue')    sh.getRange(i+1,2).setValue(b);
    if(r[0]==='tc_oficial') sh.getRange(i+1,2).setValue(o);
  }); return {ok:true};
}

// ── Marcas ───────────────────────────────────────────────────
function getMarcas() {
  return SS.getSheetByName('CONFIG')?.getDataRange().getValues()
    .filter(r=>String(r[0]).startsWith('MARCA_')).map(r=>r[1]).sort()||[];
}
function saveMarca(marca) {
  try {
    const sh = SS.getSheetByName('CONFIG'), data = sh.getDataRange().getValues();
    if (data.some(r=>String(r[0]).startsWith('MARCA_')&&String(r[1]).toUpperCase()===marca.toUpperCase()))
      return {ok:false,error:'La marca ya existe'};
    const nums = data.filter(r=>String(r[0]).startsWith('MARCA_')).map(r=>parseInt(r[0].replace('MARCA_',''))||0);
    sh.appendRow([`MARCA_${(nums.length?Math.max(...nums):0)+1}`,marca.toUpperCase(),'']);
    return {ok:true,marca:marca.toUpperCase()};
  } catch(e){ return {ok:false,error:e.message}; }
}

// ═══════════════════════════════════════════════════════════
//  INVENTARIO
// ═══════════════════════════════════════════════════════════
function getVehiculos(filtro) {
  try {
    const sh = SS.getSheetByName('VEHICULOS');
    if (!sh||sh.getLastRow()<2) return [];
    const now = new Date();
    let lista = sh.getRange(2,1,sh.getLastRow()-1,16).getValues()
      .filter(r=>r[0]&&r[6]!=='Eliminado')
      .map(r=>({id:r[0],dominio:r[1],marca:r[2],modelo:r[3],anio:r[4],tipo:r[5],estado:r[6],
        precio_toma:Number(r[7])||0,precio_tabla:Number(r[8])||0,costo_total:Number(r[9])||0,
        markup_consigna:Number(r[10])||0,ex_dueno_nombre:r[11],ex_dueno_apellido:r[12],ex_dueno_tel:r[13],
        fecha_ingreso:r[14]?_fmt(r[14],'dd/MM/yyyy'):'',observaciones:r[15],
        dias_stock:r[14]?Math.floor((now-new Date(r[14]))/864e5):0}));
    if(filtro==='stock')      lista=lista.filter(r=>r.estado==='Stock');
    else if(filtro==='reservado') lista=lista.filter(r=>r.estado==='Reservado');
    else if(filtro==='consigna')  lista=lista.filter(r=>r.tipo==='Consigna'&&r.estado!=='Vendido');
    else if(filtro==='vendidos')  lista=lista.filter(r=>r.estado==='Vendido');
    else lista=lista.filter(r=>r.estado!=='Vendido');
    return lista.sort((a,b)=>b.dias_stock-a.dias_stock);
  } catch(e){ return []; }
}
function getVehiculosDisponibles() {
  try {
    const sh = SS.getSheetByName('VEHICULOS'); if(!sh||sh.getLastRow()<2) return [];
    return sh.getRange(2,1,sh.getLastRow()-1,11).getValues()
      .filter(r=>r[0]&&(r[6]==='Stock'||r[6]==='Reservado'))
      .map(r=>({id:r[0],dominio:r[1],marca:r[2],modelo:r[3],anio:r[4],tipo:r[5],
        precio_toma:Number(r[7])||0,precio_tabla:Number(r[8])||0,
        costo_total:Number(r[9])||0,markup_consigna:Number(r[10])||0,
        label:`${r[1]} — ${r[2]} ${r[3]} ${r[4]}`}));
  } catch(e){ return []; }
}
function getTodosVehiculos() {
  try {
    const sh = SS.getSheetByName('VEHICULOS'); if(!sh||sh.getLastRow()<2) return [];
    return sh.getRange(2,1,sh.getLastRow()-1,12).getValues()
      .filter(r=>r[0]&&r[6]!=='Eliminado'&&r[6]!=='Vendido')
      .map(r=>({id:r[0],dominio:r[1],marca:r[2],modelo:r[3],anio:r[4],tipo:r[5],costo_total:Number(r[9])||0,
        ex_dueno_nombre:r[11]||'',
        label:`${r[1]} — ${r[2]} ${r[3]} ${r[4]}`}));
  } catch(e){ return []; }
}
// Solo los autos de Martín (para venta con entidad=Martin)
function getVehiculosMartin() {
  try {
    const sh = SS.getSheetByName('VEHICULOS'); if(!sh||sh.getLastRow()<2) return [];
    return sh.getRange(2,1,sh.getLastRow()-1,12).getValues()
      .filter(r=>r[0]&&r[6]!=='Eliminado'&&r[6]!=='Vendido'&&String(r[11]||'').toLowerCase().includes('martin'))
      .map(r=>({id:r[0],dominio:r[1],marca:r[2],modelo:r[3],anio:r[4],tipo:r[5],costo_total:Number(r[9])||0,
        ex_dueno_nombre:r[11]||'',
        label:`${r[1]} — ${r[2]} ${r[3]} ${r[4]}`}));
  } catch(e){ return []; }
}
function getVehiculoDetalle(id) {
  try {
    const sh = SS.getSheetByName('VEHICULOS');
    const r = sh.getRange(2,1,sh.getLastRow()-1,16).getValues().find(r=>r[0]===id);
    if (!r) return {ok:false,error:'No encontrado'};
    const vehiculo={id:r[0],dominio:r[1],marca:r[2],modelo:r[3],anio:r[4],tipo:r[5],estado:r[6],
      precio_toma:Number(r[7])||0,precio_tabla:Number(r[8])||0,costo_total:Number(r[9])||0,
      markup_consigna:Number(r[10])||0,ex_dueno_nombre:r[11],ex_dueno_apellido:r[12],ex_dueno_tel:r[13],
      fecha_ingreso:r[14]?_fmt(r[14],'dd/MM/yyyy'):'',observaciones:r[15]};
    const gSh = SS.getSheetByName('GASTOS_VEHICULOS');
    let gastos = [];
    if (gSh&&gSh.getLastRow()>1)
      gastos=gSh.getRange(2,1,gSh.getLastRow()-1,10).getValues()
        .filter(r=>r[0]&&r[1]===id)
        .map(r=>({id:r[0],id_vehiculo:r[1],dominio:r[2],fecha:r[3]?_fmt(r[3],'dd/MM/yy'):'',
          categoria:r[4],descripcion:r[5],proveedor:r[6],importe_ars:Number(r[7])||0,importe_usd:Number(r[8])||0,factura:r[9]}));
    return {ok:true,vehiculo,gastos};
  } catch(e){ return {ok:false,error:e.message}; }
}

// ── MEJORA 1: saveVehiculo con auto-caja ──────────────────────
function saveVehiculo(data) {
  try {
    const sh = SS.getSheetByName('VEHICULOS');
    let idVeh = data.id;
    if (!idVeh) {
      idVeh = _id('VEH');
      const base = Number(data.precio_toma)||Number(data.precio_compra)||0;
      sh.appendRow([idVeh,(data.dominio||'').toUpperCase().trim(),(data.marca||'').toUpperCase().trim(),
        data.modelo||'',data.anio||'',data.tipo||'Usado','Stock',base,Number(data.precio_tabla)||0,base,
        Number(data.markup_consigna)||0,data.ex_dueno_nombre||'',data.ex_dueno_apellido||'',data.ex_dueno_tel||'',
        data.fecha_ingreso?new Date(data.fecha_ingreso):new Date(),data.observaciones||'']);

      // Auto-crear egreso en caja si aplica
      if (data.registrar_caja && data.tipo !== 'Consigna' && base > 0) {
        const desc = `Compra ${(data.marca||'').toUpperCase()} ${data.modelo||''} ${data.anio||''} (${(data.dominio||'').toUpperCase()})`.trim();
        _saveCajaRaw({
          detalle: desc, concepto: 'Compra Auto', entidad: 'AGA',
          debe: 0, haber: base, medio_pago: data.medio_pago_compra || 'Efectivo',
          referencia: `VEH:${idVeh}`, observacion: '', fecha: data.fecha_ingreso || new Date()
        });
        // Si parte es en USD, registrar en DOLARES
        if (data.usd_monto && data.usd_tc) {
          SS.getSheetByName('DOLARES').appendRow([_id('USD'),
            data.fecha_ingreso?new Date(data.fecha_ingreso):new Date(),
            'Egreso', Number(data.usd_monto), Number(data.usd_tc),
            Math.round(Number(data.usd_monto)*Number(data.usd_tc)),
            data.ex_dueno_nombre||'', '', 'Compra auto']);
        }
      }
      return {ok:true,id:idVeh,nuevo:true};
    } else {
      const ids = sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
      const ri = ids.indexOf(idVeh); if(ri===-1) return {ok:false,error:'No encontrado'};
      const row = ri+2;
      [[2,(data.dominio||'').toUpperCase().trim()],[3,(data.marca||'').toUpperCase().trim()],
       [4,data.modelo||''],[5,data.anio||''],[6,data.tipo||'Usado'],
       [8,Number(data.precio_toma)||0],[9,Number(data.precio_tabla)||0],
       [11,Number(data.markup_consigna)||0],[12,data.ex_dueno_nombre||''],[13,data.ex_dueno_apellido||''],
       [14,data.ex_dueno_tel||''],[16,data.observaciones||'']]
      .forEach(([c,v])=>sh.getRange(row,c).setValue(v));
      return {ok:true,id:idVeh,nuevo:false};
    }
  } catch(e){ return {ok:false,error:e.message}; }
}

function cambiarEstadoVehiculo(id,estado) {
  try {
    const sh=SS.getSheetByName('VEHICULOS');
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    sh.getRange(ri+2,7).setValue(estado);return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}
function deleteVehiculo(id) {
  try {
    const sh=SS.getSheetByName('VEHICULOS');
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    sh.getRange(ri+2,7).setValue('Eliminado');return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}
function saveGastoVehiculo(data) {
  try {
    const sh=SS.getSheetByName('GASTOS_VEHICULOS');
    if (!data.id) {
      data.id=_id('GV');
      const vSh=SS.getSheetByName('VEHICULOS');
      const vr=vSh.getRange(2,1,vSh.getLastRow()-1,2).getValues().find(r=>r[0]===data.id_vehiculo);
      sh.appendRow([data.id,data.id_vehiculo,vr?vr[1]:'',_parseLocalDate(data.fecha),
        data.categoria||'',data.descripcion||'',data.proveedor||'',
        Number(data.importe_ars)||0,Number(data.importe_usd)||0,data.factura||'']);
    } else {
      const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
      const ri=ids.indexOf(data.id);if(ri===-1)return{ok:false};
      sh.getRange(ri+2,4).setValue(_parseLocalDate(data.fecha));
      [[5,data.categoria],[6,data.descripcion],[7,data.proveedor],
       [8,Number(data.importe_ars)||0],[9,Number(data.importe_usd)||0],[10,data.factura||'']]
      .forEach(([c,v])=>sh.getRange(ri+2,c).setValue(v));
    }
    _recalcCosto(data.id_vehiculo);
    return {ok:true,id:data.id};
  }catch(e){return{ok:false,error:e.message};}
}
function deleteGastoVehiculo(id,idVeh) {
  try {
    const sh=SS.getSheetByName('GASTOS_VEHICULOS');
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    sh.deleteRow(ri+2);_recalcCosto(idVeh);return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}
function _recalcCosto(idVeh) {
  const vSh=SS.getSheetByName('VEHICULOS');if(!vSh||vSh.getLastRow()<2)return;
  const vRows=vSh.getRange(2,1,vSh.getLastRow()-1,11).getValues();
  const vi=vRows.findIndex(r=>r[0]===idVeh);if(vi===-1)return;
  const base=Number(vRows[vi][7])||0,tipo=vRows[vi][5];
  const markup=Number(vRows[vi][10])||1.30; // markup_consigna; default 30% sobre gastos
  let gastos=0;
  const gSh=SS.getSheetByName('GASTOS_VEHICULOS');
  if(gSh&&gSh.getLastRow()>1)
    gSh.getRange(2,1,gSh.getLastRow()-1,9).getValues()
      .filter(r=>r[0]&&r[1]===idVeh).forEach(r=>{gastos+=Number(r[7])||0;});
  // Consigna: gastos × markup (interés del 30% por default). Otros: base de compra + gastos.
  vSh.getRange(vi+2,10).setValue(tipo==='Consigna' ? Math.round(gastos*markup) : base+gastos);
}

// ═══════════════════════════════════════════════════════════
//  CAJA HUB
// ═══════════════════════════════════════════════════════════

// Helper para escribir directamente en CAJA (sin sub-registros)
// Determina la sub-caja según tipo/medio_pago
function _subCajaDe(data) {
  const t = data.tipo||'';
  if (['venta_usd','compra_usd','reg_usd_ing','reg_usd_eg'].includes(t)) return 'USD';
  if (['cobro_cheque','pago_cheque'].includes(t)) return 'CHQ';
  if ((data.medio_pago||'').toLowerCase().includes('cheque')) return 'CHQ';
  return 'ARS';
}

function _saveCajaRaw(data) {
  const cSh = SS.getSheetByName('CAJA');
  const idC = _id('CAJ');
  let saldoAnt = cSh.getLastRow()>1?Number(cSh.getRange(cSh.getLastRow(),9).getValue())||0:0;
  const debe=Number(data.debe)||0,haber=Number(data.haber)||0;
  const saldo=saldoAnt+debe-haber;
  cSh.appendRow([idC,_parseLocalDate(data.fecha),data.detalle||'',
    data.observacion||'',data.concepto||'',data.entidad||'AGA',debe,haber,saldo,
    data.medio_pago||'',data.referencia||'',_subCajaDe(data)]);
  return {idCaja:idC,saldo};
}

function saveMovimientoCajaCompleto(data) {
  try {
    let idCaja, saldo;
    const isMix = data.payment_sources?.mix?.length > 0;

    if(isMix) {
      const mix = data.payment_sources.mix;
      let firstId = null;
      const cSh = SS.getSheetByName('CAJA');
      // Detectar si es ingreso (venta) por flag o por debe > haber
      const esIngreso = data.es_ingreso === true || (Number(data.debe)||0) > (Number(data.haber)||0);

      mix.forEach(comp => {
        let compAmt = 0, compTipo = '';
        if(comp.medio==='ARS')      { compAmt=comp.ars||0; compTipo=data.tipo||''; }
        else if(comp.medio==='USD') { compAmt=comp.ars||Math.round((comp.usd||0)*(comp.tc||0)); compTipo=esIngreso?'reg_usd_ing':'reg_usd_eg'; }
        else if(comp.medio==='CHQ') {
          compAmt=esIngreso?(comp.ars||0):((comp.cheques||[]).reduce((s,c)=>s+(c.importe||0),0));
          compTipo=esIngreso?'cobro_cheque':'pago_cheque';
        }
        if(!compAmt) return;

        const r=_saveCajaRaw({...data, tipo:compTipo,
          debe: esIngreso ? compAmt : 0,
          haber: esIngreso ? 0 : compAmt,
          detalle:(data.detalle||'')+(mix.length>1?` [${comp.medio}]`:'')});
        if(!firstId) firstId=r.idCaja;

        if(comp.medio==='USD'&&comp.usd&&comp.tc){
          const tipoD=esIngreso?'reg_usd_ing':'reg_usd_eg';
          _dolarInterno({tipo:tipoD,monto_usd:comp.usd,tc:comp.tc,monto_ars:compAmt,
            fecha:data.fecha,observacion:data.detalle,entidad:data.entidad||'AGA'},r.idCaja);
        }

        if(comp.medio==='CHQ'){
          if(esIngreso&&comp.cheque_data){
            // Ingreso: cheque nuevo recibido del comprador
            _chequeInterno({tipo:'cobro_cheque',
              numero:comp.cheque_data.numero||'',banco:comp.cheque_data.banco||'',
              librador:comp.cheque_data.librador||data.comprador||'',
              importe:compAmt,fecha_vto:comp.cheque_data.vto||'',
              fecha:data.fecha,entidad:data.entidad||'AGA',observacion:data.detalle},r.idCaja);
          } else if(!esIngreso&&comp.cheques){
            comp.cheques.forEach(ch=>updateEstadoCheque(ch.id,'Entregado',data.fecha));
          }
        }
      });

      idCaja=firstId;
      saldo=cSh.getLastRow()>1?Number(cSh.getRange(cSh.getLastRow(),9).getValue())||0:0;

    } else {
      // Pago simple: una sola entrada CAJA (comportamiento original)
      const cSh = SS.getSheetByName('CAJA');
      idCaja = _id('CAJ');
      let saldoAnt = cSh.getLastRow()>1?Number(cSh.getRange(cSh.getLastRow(),9).getValue())||0:0;
      const debe=Number(data.debe)||0,haber=Number(data.haber)||0;
      saldo=saldoAnt+debe-haber;
      cSh.appendRow([idCaja,_parseLocalDate(data.fecha),data.detalle||'',
        data.observacion||'',data.concepto||'',data.entidad||'AGA',debe,haber,saldo,
        data.medio_pago||'',data.referencia||'',_subCajaDe(data)]);
    }

    let extra={saldo};

    // Registros secundarios por concepto (usan el primer idCaja como referencia)
    switch(data.tipo) {
      case 'venta_auto':   extra={...extra,..._vtaInterna(data,idCaja)}; break;
      case 'gasto_auto':   extra.idSub=_gastoVehInterno(data,idCaja); break;
      case 'gasto_agencia':extra.idSub=_gastoAgInterno(data,idCaja); break;
      case 'cobro_cheque':
      case 'pago_cheque':  extra.idSub=_chequeInterno(data,idCaja); break;
      case 'venta_usd':
      case 'compra_usd':
      case 'reg_usd_ing':
      case 'reg_usd_eg':   extra.idSub=_dolarInterno(data,idCaja); break;
    }

    // Parte en USD (pago simple, no mixto)
    if(!isMix && data.usd_monto && data.usd_tc && !['venta_usd','compra_usd','reg_usd_ing','reg_usd_eg'].includes(data.tipo)) {
      _dolarInterno({tipo:data.debe>0?'reg_usd_ing':'reg_usd_eg',
        monto_usd:data.usd_monto,tc:data.usd_tc,
        monto_ars:Math.round(Number(data.usd_monto)*Number(data.usd_tc)),
        contraparte:data.comprador||data.proveedor||'',
        fecha:data.fecha,observacion:'Parte en USD — '+data.detalle,entidad:data.entidad||'AGA'},idCaja);
    }

    // payment_sources método único (USD o CHQ)
    if(!isMix && data.payment_sources) {
      const ps=data.payment_sources;
      if(ps.usd&&ps.usd.monto_usd&&ps.usd.tc){
        const esIngreso=data.es_ingreso===true||(Number(data.debe)||0)>(Number(data.haber)||0);
        _dolarInterno({tipo:esIngreso?'reg_usd_ing':'reg_usd_eg',monto_usd:ps.usd.monto_usd,tc:ps.usd.tc,
          monto_ars:Math.round(ps.usd.monto_usd*ps.usd.tc),
          fecha:data.fecha,observacion:'Cobro en USD — '+data.detalle,entidad:data.entidad||'AGA'},idCaja);
      }
      if(ps.cheques&&ps.cheques.length)
        ps.cheques.forEach(ch=>updateEstadoCheque(ch.id,'Entregado',data.fecha));
      // Cheque nuevo recibido (cobro de ingreso)
      if(ps.chq_nuevo&&ps.chq_nuevo.importe>0){
        _chequeInterno({tipo:'cobro_cheque',
          numero:ps.chq_nuevo.numero||'',banco:ps.chq_nuevo.banco||'',
          librador:data.comprador||'',importe:ps.chq_nuevo.importe,
          fecha_vto:ps.chq_nuevo.vto||'',
          fecha:data.fecha,entidad:data.entidad||'AGA',observacion:data.detalle},idCaja);
      }
    }

    return {ok:true,idCaja,...extra};
  }catch(e){return{ok:false,error:e.message};}
}

function registrarVenta(data) {
  try {
    const vSh=SS.getSheetByName('VEHICULOS');
    const vRows=vSh.getRange(2,1,vSh.getLastRow()-1,16).getValues();
    const vi=vRows.findIndex(r=>r[0]===data.id_vehiculo);
    if(vi===-1)return{ok:false,error:'Vehículo no encontrado'};
    const vr=vRows[vi];
    const costo=Number(vr[9])||0,precio=Number(data.precio_venta)||0;
    const esConsigna=vr[5]==='Consigna';
    // Consigna: ganancia = 5% comisión + markup ya incluido en costo (30% sobre gastos)
    const ganancia=esConsigna ? Math.round(precio*0.05) : precio-costo;
    const desc=`${vr[2]} ${vr[3]} ${vr[4]}`.trim();
    const idVenta=_id('VTA');
    SS.getSheetByName('VENTAS').appendRow([idVenta,data.id_vehiculo,vr[1],desc,
      data.fecha_venta?new Date(data.fecha_venta):new Date(),
      data.comprador||'',data.vendedor||'',precio,
      Number(data.efectivo)||0,Number(data.vehiculo_entrega)||0,
      Number(data.financiacion)||0,Number(data.tc_dia)||0,costo,ganancia,data.observaciones||'']);
    vSh.getRange(vi+2,7).setValue('Vendido');
    if(data.comprador)_upsertPersona(data.comprador,'Comprador');
    // Auto-crear en caja (respetando entidad)
    const entidad=data.entidad||'AGA';
    const cajaRes=_saveCajaRaw({detalle:`Venta ${desc} (${vr[1]})`,concepto:`Venta Auto ${vr[5]}`,entidad,
      tipo:'venta_auto_directo',debe:precio,haber:0,medio_pago:data.medio_pago||'Efectivo',
      referencia:`VTA:${idVenta}`,observacion:data.observaciones||'',fecha:data.fecha_venta||new Date()});
    return {ok:true,idVenta,ganancia,vehiculoStr:desc,esConsigna,...cajaRes};
  }catch(e){return{ok:false,error:e.message};}
}

function _vtaInterna(data,idCaja) {
  const vSh=SS.getSheetByName('VEHICULOS');
  const vRows=vSh.getRange(2,1,vSh.getLastRow()-1,16).getValues();
  const vi=vRows.findIndex(r=>r[0]===data.id_vehiculo);
  if(vi===-1)return{ok:false,error:'Vehículo no encontrado: '+data.id_vehiculo};
  const vr=vRows[vi];
  const costo=Number(vr[9])||0,precio=Number(data.precio_venta)||Number(data.debe)||0;
  const esConsigna=vr[5]==='Consigna';
  // Consigna: comisión editable (default 5%)
  const comisionPct=esConsigna?(Number(data.comision_pct)||5):100;
  const ganancia=esConsigna ? Math.round(precio*comisionPct/100) : precio-costo;
  const desc=`${vr[2]} ${vr[3]} ${vr[4]}`.trim(),idVenta=_id('VTA');
  SS.getSheetByName('VENTAS').appendRow([idVenta,data.id_vehiculo,vr[1],desc,
    _parseLocalDate(data.fecha),
    data.comprador||'',data.vendedor||'',precio,
    Number(data.efectivo)||0,Number(data.vehiculo_entrega)||0,
    Number(data.financiacion)||0,Number(data.tc_dia)||0,costo,ganancia,data.observacion||'']);
  vSh.getRange(vi+2,7).setValue('Vendido');
  SS.getSheetByName('CAJA').getRange(SS.getSheetByName('CAJA').getLastRow(),11).setValue(`VTA:${idVenta}`);
  if(data.comprador)_upsertPersona(data.comprador,'Comprador');
  return{ok:true,idVenta,ganancia,vehiculoStr:desc,esConsigna,comisionPct};
}
function _gastoVehInterno(data,idCaja) {
  const id=_id('GV');
  const vSh=SS.getSheetByName('VEHICULOS');
  const vr=vSh.getRange(2,1,vSh.getLastRow()-1,2).getValues().find(r=>r[0]===data.id_vehiculo);
  SS.getSheetByName('GASTOS_VEHICULOS').appendRow([id,data.id_vehiculo,vr?vr[1]:'',
    _parseLocalDate(data.fecha),data.categoria||'',
    data.descripcion||'',data.proveedor||'',Number(data.haber)||0,0,'']);
  _recalcCosto(data.id_vehiculo);
  SS.getSheetByName('CAJA').getRange(SS.getSheetByName('CAJA').getLastRow(),11).setValue(`GV:${id}`);
  return id;
}
function _gastoAgInterno(data,idCaja) {
  const id=_id('GA');
  const neto=Number(data.neto)||Number(data.haber)||0,iva=Number(data.iva)||0;
  const total=Number(data.total)||neto+iva||Number(data.haber)||0;
  SS.getSheetByName('GASTOS_AGENCIA').appendRow([id,_parseLocalDate(data.fecha),
    _periodo(data.fecha),data.categoria||'',data.tipo||'Variable',
    data.proveedor||'',data.descripcion||'',neto,iva,total,
    data.canal||1,data.medio_pago||'',idCaja,data.observacion||'']);
  SS.getSheetByName('CAJA').getRange(SS.getSheetByName('CAJA').getLastRow(),11).setValue(`GA:${id}`);
  return id;
}
function _chequeInterno(data,idCaja) {
  const id=_id('CHQ');
  const esC=data.tipo==='cobro_cheque';
  const entidad=data.entidad||'AGA';
  // Schema v2 (13 cols): ID|NUMERO|TIPO|BANCO|LIBRADOR|BENEFICIARIO|IMPORTE|FECHA_VTO|FECHA_COBRO|ESTADO|ID_CAJA_REF|OBSERVACIONES|ENTIDAD
  SS.getSheetByName('CHEQUES').appendRow([id,data.numero||'',esC?'De tercero':'Propio',
    data.banco||'',data.librador||'',data.beneficiario||data.librador||'',
    Number(data.importe)||Number(data.debe)||Number(data.haber)||0,
    data.fecha_vto?new Date(data.fecha_vto):'',
    '', // FECHA_COBRO vacío hasta que se cobre
    esC?'En cartera':'Emitido', idCaja, data.observacion||'', entidad]);
  SS.getSheetByName('CAJA').getRange(SS.getSheetByName('CAJA').getLastRow(),11).setValue(`CHQ:${id}`);
  return id;
}
// DOLARES schema (10 cols): ID|FECHA|TIPO|MONTO_USD|TC|MONTO_ARS|CONTRAPARTE|ID_CAJA_REF|OBSERVACIONES|ENTIDAD
function _dolarInterno(data,idCaja) {
  const id=_id('USD');
  const tipoMap={venta_usd:'Venta',compra_usd:'Compra',reg_usd_ing:'Ingreso',reg_usd_eg:'Egreso'};
  const tipo=tipoMap[data.tipo]||data.tipo||'Ingreso';
  const monto_usd=Number(data.monto_usd)||0,tc=Number(data.tc)||Number(data.tc_dia)||0;
  const monto_ars=Number(data.monto_ars)||Math.round(monto_usd*tc)||Number(data.debe)||Number(data.haber)||0;
  const entidad=data.entidad||'AGA';
  SS.getSheetByName('DOLARES').appendRow([id,_parseLocalDate(data.fecha),
    tipo,monto_usd,tc,monto_ars,data.contraparte||'',idCaja,data.observacion||'',entidad]);
  if(idCaja) SS.getSheetByName('CAJA').getRange(SS.getSheetByName('CAJA').getLastRow(),11).setValue(`USD:${id}`);
  return id;
}

// ── Consultas caja ────────────────────────────────────────────
function getMovimientosCaja(mes,anio,subCaja,entidad) {
  try {
    const sh=SS.getSheetByName('CAJA');if(!sh||sh.getLastRow()<2)return{ok:true,movimientos:[],saldo:0};
    const cols=Math.min(sh.getLastColumn(),12);
    const rows=sh.getRange(2,1,sh.getLastRow()-1,cols).getValues().filter(r=>r[0]);
    let filtrados=rows;
    const ent=entidad||'AGA';
    filtrados=filtrados.filter(r=>r[5]===ent||(!r[5]&&ent==='AGA'));
    if(mes&&anio)filtrados=filtrados.filter(r=>{const f=new Date(r[1]);return f.getMonth()===parseInt(mes)-1&&f.getFullYear()===parseInt(anio);});
    if(subCaja&&subCaja!=='general')filtrados=filtrados.filter(r=>(r[11]||'ARS')===subCaja);
    const movimientos=filtrados.map(r=>({id:r[0],fecha:r[1]?_fmt(r[1],'dd/MM/yyyy'):'',detalle:r[2],
      observacion:r[3],concepto:r[4],entidad:r[5],debe:Number(r[6])||0,haber:Number(r[7])||0,
      saldo:Number(r[8])||0,medio_pago:r[9],referencia:r[10]||'',sub_caja:r[11]||'ARS'}));
    // Saldo = running total for this entity
    let saldoEnt=0;
    rows.filter(r=>r[5]===ent||(!r[5]&&ent==='AGA')).forEach(r=>{saldoEnt+=(Number(r[6])||0)-(Number(r[7])||0);});
    return{ok:true,movimientos,saldo:saldoEnt};
  }catch(e){return{ok:false,error:e.message};}
}
function deleteMovimientoCaja(id) {
  try {
    const sh=SS.getSheetByName('CAJA');if(sh.getLastRow()<2)return{ok:false};
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    sh.deleteRow(ri+2);_recalcSaldoCaja();return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}
function _recalcSaldoCaja() {
  const sh=SS.getSheetByName('CAJA');if(!sh||sh.getLastRow()<2)return;
  let saldo=0;
  sh.getRange(2,1,sh.getLastRow()-1,9).getValues().forEach((r,i)=>{
    if(!r[0])return;saldo+=(Number(r[6])||0)-(Number(r[7])||0);sh.getRange(i+2,9).setValue(saldo);
  });
}

// ── Gastos agencia ────────────────────────────────────────────
function getGastosAgencia(mes,anio) {
  try {
    const sh=SS.getSheetByName('GASTOS_AGENCIA');if(!sh||sh.getLastRow()<2)return[];
    let rows=sh.getRange(2,1,sh.getLastRow()-1,14).getValues().filter(r=>r[0]);
    if(mes&&anio)rows=rows.filter(r=>{const f=new Date(r[1]);return f.getMonth()===parseInt(mes)-1&&f.getFullYear()===parseInt(anio);});
    return rows.map(r=>({id:r[0],fecha:r[1]?_fmt(r[1],'dd/MM/yyyy'):'',periodo:r[2],
      categoria:r[3],tipo:r[4],proveedor:r[5],descripcion:r[6],
      neto:Number(r[7])||0,iva:Number(r[8])||0,total:Number(r[9])||0,
      canal:r[10],medio_pago:r[11],id_caja_ref:r[12],observaciones:r[13]}))
      .sort((a,b)=>new Date(b.fecha)-new Date(a.fecha));
  }catch(e){return[];}
}

// ── Cheques ────────────────────────────────────────────────────

// CHEQUES schema v2 (13 cols): ID|NUMERO|TIPO|BANCO|LIBRADOR|BENEFICIARIO|IMPORTE|FECHA_VTO|FECHA_COBRO|ESTADO|ID_CAJA_REF|OBSERVACIONES|ENTIDAD
function getCheques(filtro, entidad) {
  try {
    const sh=SS.getSheetByName('CHEQUES');if(!sh||sh.getLastRow()<2)return[];
    const now=new Date(), ent=entidad||'AGA';
    const ncols=Math.min(sh.getLastColumn(),13);
    let rows=sh.getRange(2,1,sh.getLastRow()-1,ncols).getValues().filter(r=>r[0]);
    // col 13 = índice 12 = ENTIDAD (default 'AGA' para filas anteriores)
    rows=rows.filter(r=>(r[12]||'AGA')===ent);
    if(filtro==='en_cartera')    rows=rows.filter(r=>r[9]==='En cartera');
    else if(filtro==='depositados') rows=rows.filter(r=>r[9]==='Depositado'||r[9]==='Acreditado');
    else if(filtro==='propios')  rows=rows.filter(r=>r[2]==='Propio');
    return rows.map(r=>({id:r[0],numero:r[1],tipo:r[2],banco:r[3],librador:r[4],beneficiario:r[5],
      importe:Number(r[6])||0,
      fecha_vto:r[7]?_fmt(r[7],'dd/MM/yyyy'):'',
      fecha_cobro:r[8]?_fmt(r[8],'dd/MM/yyyy'):'',
      estado:r[9],id_caja_ref:r[10],observaciones:r[11],entidad:r[12]||'AGA',
      dias_vto:r[7]?Math.ceil((new Date(r[7])-now)/864e5):null}))
      .sort((a,b)=>(a.dias_vto??9999)-(b.dias_vto??9999));
  }catch(e){return[];}
}

function updateEstadoCheque(id,estado,fechaCobro) {
  try {
    const sh=SS.getSheetByName('CHEQUES');
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    const row=ri+2;
    sh.getRange(row,10).setValue(estado);
    if(fechaCobro&&(estado==='Acreditado'||estado==='Cobrado'||estado==='Depositado'))
      sh.getRange(row,9).setValue(new Date(fechaCobro));
    return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}

function updateCheque(id,data) {
  try {
    const sh=SS.getSheetByName('CHEQUES');
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    const row=ri+2;
    if(data.numero!==undefined)  sh.getRange(row,2).setValue(data.numero);
    if(data.banco!==undefined)   sh.getRange(row,4).setValue(data.banco);
    if(data.librador!==undefined)sh.getRange(row,5).setValue(data.librador);
    if(data.beneficiario!==undefined)sh.getRange(row,6).setValue(data.beneficiario);
    if(data.importe!==undefined) sh.getRange(row,7).setValue(Number(data.importe)||0);
    if(data.fecha_vto!==undefined)sh.getRange(row,8).setValue(data.fecha_vto?new Date(data.fecha_vto):'');
    if(data.fecha_cobro!==undefined)sh.getRange(row,9).setValue(data.fecha_cobro?new Date(data.fecha_cobro):'');
    if(data.estado!==undefined)  sh.getRange(row,10).setValue(data.estado);
    if(data.observaciones!==undefined)sh.getRange(row,12).setValue(data.observaciones);
    return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}

function deleteCheque(id) {
  try {
    const sh=SS.getSheetByName('CHEQUES');
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    sh.deleteRow(ri+2);return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}

function getChequesDisponibles(entidad) {
  try {
    const sh=SS.getSheetByName('CHEQUES');if(!sh||sh.getLastRow()<2)return[];
    const now=new Date(), ent=entidad||'AGA';
    const ncols=Math.min(sh.getLastColumn(),13);
    return sh.getRange(2,1,sh.getLastRow()-1,ncols).getValues()
      .filter(r=>r[0]&&r[9]==='En cartera'&&r[2]==='De tercero'&&(r[12]||'AGA')===ent)
      .map(r=>({id:r[0],numero:r[1],banco:r[3],librador:r[4],importe:Number(r[6])||0,
        fecha_vto:r[7]?_fmt(r[7],'dd/MM/yyyy'):'',
        dias_vto:r[7]?Math.ceil((new Date(r[7])-now)/864e5):null}));
  }catch(e){return[];}
}

function deleteGastoAgencia(id) {
  try {
    const sh=SS.getSheetByName('GASTOS_AGENCIA');
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    sh.deleteRow(ri+2);return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}

function deleteDolar(id) {
  try {
    const sh=SS.getSheetByName('DOLARES');
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    sh.deleteRow(ri+2);return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}

function getSaldosSubCaja(entidad) {
  try {
    const ent = entidad||'AGA';
    const cSh=SS.getSheetByName('CAJA');
    let saldoARS=0,saldoUSD=0,saldoCHQ=0,saldoGeneral=0;
    if(cSh&&cSh.getLastRow()>1){
      const cols=Math.min(cSh.getLastColumn(),12);
      cSh.getRange(2,1,cSh.getLastRow()-1,cols).getValues().filter(r=>r[0]).forEach(r=>{
        if(r[5]!==ent) return;
        const debe=Number(r[6])||0,haber=Number(r[7])||0,sub=r[11]||'ARS',neto=debe-haber;
        saldoGeneral+=neto;
        if(sub==='USD') saldoUSD+=neto; else if(sub==='CHQ') saldoCHQ+=neto; else saldoARS+=neto;
      });
    }
    // Posicion USD: ahora filtrada por entidad (col 10 = índice 9)
    let posUSD=0;
    const dSh=SS.getSheetByName('DOLARES');
    if(dSh&&dSh.getLastRow()>1){
      const dCols=Math.min(dSh.getLastColumn(),10);
      dSh.getRange(2,1,dSh.getLastRow()-1,dCols).getValues().filter(r=>r[0]).forEach(r=>{
        if((r[9]||'AGA')!==ent) return; // filtrar por entidad
        const t=r[2];
        if(t==='Compra'||t==='Ingreso'||t==='Cobro')posUSD+=Number(r[3])||0;
        if(t==='Venta'||t==='Egreso'||t==='Pago')    posUSD-=Number(r[3])||0;
      });
    }
    // Cheques en cartera: ahora filtrados por entidad (col 13 = índice 12)
    let chequesEnCartera=0;
    const chSh=SS.getSheetByName('CHEQUES');
    if(chSh&&chSh.getLastRow()>1){
      const chCols=Math.min(chSh.getLastColumn(),13);
      chSh.getRange(2,1,chSh.getLastRow()-1,chCols).getValues()
        .filter(r=>r[0]&&r[9]==='En cartera'&&r[2]==='De tercero'&&(r[12]||'AGA')===ent)
        .forEach(r=>{chequesEnCartera+=Number(r[6])||0;});
    }
    const listas=getConfigListas();
    const tcBlue=listas.tc_blue||1450;
    return{ok:true,saldoARS,saldoUSD,saldoCHQ,saldoGeneral,posUSD,chequesEnCartera,tcBlue,
      saldoUSD_ars:Math.round(posUSD*tcBlue)};
  }catch(e){return{ok:false,error:e.message};}
}


function getMovimientoDetalle(idCaja) {
  try {
    const cSh=SS.getSheetByName('CAJA');if(!cSh||cSh.getLastRow()<2)return{ok:false};
    const cols=Math.min(cSh.getLastColumn(),12);
    const rows=cSh.getRange(2,1,cSh.getLastRow()-1,cols).getValues();
    // Buscar entrada principal Y entradas del mismo concepto (para pagos mixtos)
    const main=rows.find(r=>r[0]===idCaja);
    if(!main)return{ok:false,error:'No encontrado'};
    const entry={id:main[0],fecha:main[1]?_fmt(main[1],'dd/MM/yyyy HH:mm'):'',
      detalle:main[2],observacion:main[3],concepto:main[4],entidad:main[5],
      debe:Number(main[6])||0,haber:Number(main[7])||0,saldo:Number(main[8])||0,
      medio_pago:main[9],referencia:main[10]||'',sub_caja:main[11]||'ARS'};
    // Buscar registros relacionados (misma referencia o mismo detalle base)
    const ref=main[10]||'';
    let linked=null;
    if(ref){
      const parts=ref.split(':');const tipo=parts[0],id2=parts.slice(1).join(':');
      if(tipo==='VTA'&&id2){
        const vtSh=SS.getSheetByName('VENTAS');
        if(vtSh){const vr=vtSh.getRange(2,1,vtSh.getLastRow()-1,15).getValues().find(r=>r[0]===id2);
        if(vr)linked={tipo:'Venta',datos:{vehiculo:vr[3],comprador:vr[5],precio:Number(vr[7]),ganancia:Number(vr[13]),costo:Number(vr[12]),vendedor:vr[6]}};}
      }
      if(tipo==='GA'&&id2){
        const gaSh=SS.getSheetByName('GASTOS_AGENCIA');
        if(gaSh){const gr=gaSh.getRange(2,1,gaSh.getLastRow()-1,14).getValues().find(r=>r[0]===id2);
        if(gr)linked={tipo:'Gasto Agencia',datos:{categoria:gr[3],proveedor:gr[5],descripcion:gr[6],neto:Number(gr[7]),iva:Number(gr[8]),total:Number(gr[9])}};}
      }
      if(tipo==='GV'&&id2){
        const gvSh=SS.getSheetByName('GASTOS_VEHICULOS');
        if(gvSh){const gv=gvSh.getRange(2,1,gvSh.getLastRow()-1,10).getValues().find(r=>r[0]===id2);
        if(gv)linked={tipo:'Gasto Vehículo',datos:{vehiculo:gv[2],categoria:gv[4],descripcion:gv[5],proveedor:gv[6],importe:Number(gv[7])}};}
      }
      if(tipo==='CHQ'&&id2){
        const chSh=SS.getSheetByName('CHEQUES');
        if(chSh){const ch=chSh.getRange(2,1,chSh.getLastRow()-1,12).getValues().find(r=>r[0]===id2);
        if(ch)linked={tipo:'Cheque',datos:{numero:ch[1],banco:ch[3],librador:ch[4],importe:Number(ch[6]),fecha_vto:ch[7]?_fmt(ch[7],'dd/MM/yyyy'):'',estado:ch[9]}};}
      }
      if(tipo==='USD'&&id2){
        const dSh=SS.getSheetByName('DOLARES');
        if(dSh){const dr=dSh.getRange(2,1,dSh.getLastRow()-1,9).getValues().find(r=>r[0]===id2);
        if(dr)linked={tipo:'USD',datos:{tipo_mov:dr[2],monto_usd:Number(dr[3]),tc:Number(dr[4]),monto_ars:Number(dr[5]),contraparte:dr[6]}};}
      }
      if(tipo==='VEH'&&id2){
        const vhSh=SS.getSheetByName('VEHICULOS');
        if(vhSh){const vh=vhSh.getRange(2,1,vhSh.getLastRow()-1,11).getValues().find(r=>r[0]===id2);
        if(vh)linked={tipo:'Vehículo',datos:{dominio:vh[1],marca:vh[2],modelo:vh[3],anio:vh[4],tipo:vh[5],estado:vh[6],precio:Number(vh[7])}};}
      }
    }
    return{ok:true,entry,linked};
  }catch(e){return{ok:false,error:e.message};}
}


function getDolares(mes,anio,entidad) {
  try {
    const sh=SS.getSheetByName('DOLARES');
    if(!sh||sh.getLastRow()<2) return [];
    const ent=entidad||'AGA';
    const lc=sh.getLastColumn();
    let rows=sh.getRange(2,1,sh.getLastRow()-1,lc).getValues().filter(r=>r[0]);
    rows=rows.filter(r=>(r[9]||'AGA')===ent);
    if(mes&&anio){
      const m=parseInt(mes)-1,y=parseInt(anio);
      rows=rows.filter(r=>{if(!r[1])return false;const f=new Date(r[1]);return f.getMonth()===m&&f.getFullYear()===y;});
    }
    // Ordenar ANTES del map mientras r[1] es Date object (evita new Date('dd/MM/yyyy') inválido)
    rows.sort((a,b)=>{const ta=a[1]?new Date(a[1]).getTime():0,tb=b[1]?new Date(b[1]).getTime():0;return tb-ta;});
    return rows.map(r=>({id:r[0],fecha:r[1]?_fmt(r[1],'dd/MM/yyyy'):'',tipo:r[2]||'',
      monto_usd:Number(r[3])||0,tc:Number(r[4])||0,monto_ars:Number(r[5])||0,
      contraparte:r[6]||'',id_caja_ref:r[7]||'',observaciones:r[8]||'',entidad:r[9]||'AGA'}));
  }catch(e){Logger.log('getDolares error: '+e.toString());return[];}
}

// ── Ventas ─────────────────────────────────────────────────────
function getVentas(lim) {
  try {
    const sh=SS.getSheetByName('VENTAS');if(!sh||sh.getLastRow()<2)return[];
    return sh.getRange(2,1,sh.getLastRow()-1,15).getValues()
      .filter(r=>r[0]).sort((a,b)=>new Date(b[4])-new Date(a[4])).slice(0,lim||50)
      .map(r=>({id:r[0],id_vehiculo:r[1],dominio:r[2],vehiculo:r[3],
        fecha:r[4]?_fmt(r[4],'dd/MM/yyyy'):'',comprador:r[5],vendedor:r[6],
        precio_venta:Number(r[7])||0,costo_total:Number(r[12])||0,ganancia:Number(r[13])||0,observaciones:r[14]}));
  }catch(e){return[];}
}

// ═══════════════════════════════════════════════════════════
//  PASO 6 — REPORTES / ESTADO DE RESULTADOS
// ═══════════════════════════════════════════════════════════
function getReporteMensual(mes, anio) {
  try {
    const m=parseInt(mes||new Date().getMonth()+1);
    const a=parseInt(anio||new Date().getFullYear());
    const start=new Date(a,m-1,1),end=new Date(a,m,1);

    // Ventas del mes
    const vtSh=SS.getSheetByName('VENTAS');
    let ventas=[];
    if(vtSh&&vtSh.getLastRow()>1)
      ventas=vtSh.getRange(2,1,vtSh.getLastRow()-1,15).getValues()
        .filter(r=>r[0]&&new Date(r[4])>=start&&new Date(r[4])<end);

    const totalVentas=ventas.reduce((s,r)=>s+(Number(r[7])||0),0);
    const totalCostos=ventas.reduce((s,r)=>s+(Number(r[12])||0),0);
    const totalGanVentas=ventas.reduce((s,r)=>s+(Number(r[13])||0),0);

    // Por tipo de vehículo (join con VEHICULOS para obtener el tipo real)
    const porTipo={};
    const vehTipoMap={};
    if(vSh&&vSh.getLastRow()>1)
      vSh.getRange(2,1,vSh.getLastRow()-1,6).getValues().forEach(r=>{if(r[0])vehTipoMap[r[0]]=r[5]||'Usado';});
    ventas.forEach(r=>{
      const tipo=vehTipoMap[r[1]]||'Usado';
      if(!porTipo[tipo])porTipo[tipo]={count:0,total:0,ganancia:0};
      porTipo[tipo].count++; porTipo[tipo].total+=Number(r[7])||0; porTipo[tipo].ganancia+=Number(r[13])||0;
    });

    // Gastos agencia del mes
    const gaSh=SS.getSheetByName('GASTOS_AGENCIA');
    let gastosAg=[];
    if(gaSh&&gaSh.getLastRow()>1)
      gastosAg=gaSh.getRange(2,1,gaSh.getLastRow()-1,14).getValues()
        .filter(r=>r[0]&&new Date(r[1])>=start&&new Date(r[1])<end);

    const totalGastosAg=gastosAg.reduce((s,r)=>s+(Number(r[9])||0),0);
    const porCategoria={};
    gastosAg.forEach(r=>{const cat=r[3]||'Sin categoría';porCategoria[cat]=(porCategoria[cat]||0)+(Number(r[9])||0);});

    // Top ventas del mes
    const topVentas=ventas.sort((a,b)=>(Number(b[13])-Number(a[13]))).slice(0,8)
      .map(r=>({vehiculo:r[3],precio:Number(r[7]),costo:Number(r[12]),ganancia:Number(r[13]),
        margen:Number(r[12])>0?((Number(r[13])/Number(r[12]))*100).toFixed(1):0,
        fecha:r[4]?_fmt(r[4],'dd/MM'):''}));

    // Saldo caja actual
    const cSh=SS.getSheetByName('CAJA');
    const saldoCaja=cSh&&cSh.getLastRow()>1?Number(cSh.getRange(cSh.getLastRow(),9).getValue())||0:0;

    // Posicion USD
    const dSh=SS.getSheetByName('DOLARES');
    let posUSD=0;
    if(dSh&&dSh.getLastRow()>1)
      dSh.getRange(2,1,dSh.getLastRow()-1,6).getValues().filter(r=>r[0]).forEach(r=>{
        const t=r[2];
        if(t==='Compra'||t==='Ingreso')posUSD+=Number(r[3])||0;
        if(t==='Venta'||t==='Egreso')  posUSD-=Number(r[3])||0;
      });

    // Stock value y rotación
    const vSh=SS.getSheetByName('VEHICULOS');
    let stockValue=0,stockCount=0;
    // Mapa vehículo → fecha_ingreso para rotación
    const vFechaMap={};
    if(vSh&&vSh.getLastRow()>1){
      vSh.getRange(2,1,vSh.getLastRow()-1,16).getValues().filter(r=>r[0]).forEach(r=>{
        if(r[6]==='Stock'||r[6]==='Reservado'){stockValue+=Number(r[9])||0;stockCount++;}
        if(r[14]) vFechaMap[r[0]]=new Date(r[14]);
      });
    }

    // KPI: Rotación de stock = promedio días en stock de autos vendidos este mes
    let totalDiasStock=0,countRotacion=0;
    ventas.forEach(r=>{
      const fi=vFechaMap[r[1]];
      if(fi){const dias=Math.floor((new Date(r[4])-fi)/864e5);if(dias>=0){totalDiasStock+=dias;countRotacion++;}}
    });
    const rotacionDias=countRotacion>0?Math.round(totalDiasStock/countRotacion):null;

    // KPI: Costo de comercialización por auto = gastos agencia / autos vendidos
    const costoComercializacion=ventas.length>0?Math.round(totalGastosAg/ventas.length):null;

    const periodo=`${String(m).padStart(2,'0')}/${a}`;
    return {ok:true,mes:m,anio:a,periodo,
      ventas:{count:ventas.length,total:totalVentas,costos:totalCostos,ganancia:totalGanVentas,
        margen:totalCostos>0?((totalGanVentas/totalCostos)*100).toFixed(1):0,
        promPrecio:ventas.length?Math.round(totalVentas/ventas.length):0,
        lista:topVentas},
      gastosAgencia:{total:totalGastosAg,count:gastosAg.length,porCategoria},
      resultado:totalGanVentas-totalGastosAg,
      saldoCaja,posUSD,stockValue,stockCount,
      rotacionDias,costoComercializacion};
  }catch(e){return{ok:false,error:e.message};}
}

// ═══════════════════════════════════════════════════════════
//  PASO 7 — PERSONAS / CONTACTOS
// ═══════════════════════════════════════════════════════════
function getPersonas(tipo) {
  try {
    const sh=SS.getSheetByName('PERSONAS');if(!sh||sh.getLastRow()<2)return[];
    let rows=sh.getRange(2,1,sh.getLastRow()-1,8).getValues().filter(r=>r[0]);
    if(tipo&&tipo!=='todos')rows=rows.filter(r=>r[3]===tipo);
    return rows.map(r=>({id:r[0],nombre:r[1],apellido:r[2],tipo:r[3],telefono:r[4],dni:r[5],email:r[6],
      observaciones:r[7],nombre_completo:`${r[1]} ${r[2]}`.trim()}))
      .sort((a,b)=>a.nombre_completo.localeCompare(b.nombre_completo));
  }catch(e){return[];}
}
function getVendedores(){return getPersonas('Vendedor');}
function buscarPersonas(q) {
  try {
    const sh=SS.getSheetByName('PERSONAS');if(!sh||sh.getLastRow()<2)return[];
    const query=(q||'').toLowerCase();
    return sh.getRange(2,1,sh.getLastRow()-1,3).getValues()
      .filter(r=>r[0]&&(`${r[1]} ${r[2]}`).toLowerCase().includes(query))
      .slice(0,8).map(r=>`${r[1]} ${r[2]}`.trim());
  }catch(e){return[];}
}
function savePersona(data) {
  try {
    const sh=SS.getSheetByName('PERSONAS');
    if(!data.id){
      data.id=_id('PER');
      sh.appendRow([data.id,data.nombre||'',data.apellido||'',data.tipo||'Comprador',
        data.telefono||'',data.dni||'',data.email||'',data.observaciones||'']);
    } else {
      const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
      const ri=ids.indexOf(data.id);if(ri===-1)return{ok:false};
      [data.nombre,data.apellido,data.tipo,data.telefono,data.dni,data.email,data.observaciones]
        .forEach((v,i)=>sh.getRange(ri+2,i+2).setValue(v||''));
    }
    return{ok:true,id:data.id};
  }catch(e){return{ok:false,error:e.message};}
}
function deletePersona(id) {
  try {
    const sh=SS.getSheetByName('PERSONAS');
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
    const ri=ids.indexOf(id);if(ri===-1)return{ok:false};
    sh.deleteRow(ri+2);return{ok:true};
  }catch(e){return{ok:false,error:e.message};}
}
function _upsertPersona(nombre,tipo) {
  const sh=SS.getSheetByName('PERSONAS');if(!sh)return;
  const ex=sh.getLastRow()>1?sh.getRange(2,2,sh.getLastRow()-1,2).getValues().map(r=>`${r[0]} ${r[1]}`.trim().toLowerCase()):[];
  if(!ex.includes(nombre.toLowerCase())){const p=nombre.split(' ');sh.appendRow([_id('PER'),p[0],p.slice(1).join(' '),tipo,'','','','']);}
}

// ── Config UI ─────────────────────────────────────────────────
function getConfigUI() {try{return{};} catch(e){return{};}}
function saveConfigUI(prefs) {
  try {
    const sh=SS.getSheetByName('CONFIG'),data=sh.getDataRange().getValues();
    Object.entries(prefs).forEach(([key,val])=>{
      const ri=data.findIndex(r=>r[0]===key);
      if(ri!==-1)sh.getRange(ri+1,2).setValue(val);
      else sh.appendRow([key,val,'Configuración UI']);
    });
    return{ok:true};
  }catch(e){return{ok:false};}
}

// ── Utils ────────────────────────────────────────────────────
function _id(p){return`${p}-${Date.now()}-${Math.floor(Math.random()*999)}`;}
function _fmt(d,fmt){try{return d?Utilities.formatDate(new Date(d),'America/Argentina/Buenos_Aires',fmt):''}catch(e){return'';}}
function _periodo(f){const d=f?new Date(f):new Date();return Utilities.formatDate(d,'America/Argentina/Buenos_Aires','MM/yyyy');}
// Parsea 'YYYY-MM-DD' al mediodía local para evitar desfase de zona horaria UTC-3
function _parseLocalDate(v){
  if(!v)return new Date();
  if(v instanceof Date)return v;
  const p=String(v).split('-');
  if(p.length>=3)return new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]),12,0,0);
  return new Date(v);
}

// ── Setup ─────────────────────────────────────────────────────
function setupSheets() {
  const ui=SpreadsheetApp.getUi();
  if(ui.alert('⚠️ Setup AGA v2.1','Borra y recrea todas las hojas. Solo usar la 1era vez.\n\n¿Continuar?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  const schemas=[
    {name:'CONFIG',color:'#1a1a2e',headers:['CLAVE','VALOR','DESCRIPCION'],data:[
      ['tc_blue',1450,'Dólar blue'],['tc_oficial',1100,'Dólar oficial'],['markup_consigna_default',1.30,''],
      ['secretaria@gmail.com','admin','→ Reemplazar'],['dueno@gmail.com','dueno','→ Reemplazar'],['vendedor@gmail.com','vendedor','→ Reemplazar'],['martin@gmail.com','admin','→ Email de Martín (reemplazar)'],
      ['MARCA_1','VOLKSWAGEN',''],['MARCA_2','FORD',''],['MARCA_3','CHEVROLET',''],['MARCA_4','TOYOTA',''],['MARCA_5','PEUGEOT',''],
      ['MARCA_6','RENAULT',''],['MARCA_7','FIAT',''],['MARCA_8','HONDA',''],['MARCA_9','CITROEN',''],['MARCA_10','HYUNDAI',''],
      ['MARCA_11','KIA',''],['MARCA_12','NISSAN',''],['MARCA_13','MERCEDES-BENZ',''],['MARCA_14','BMW',''],['MARCA_15','AUDI',''],
      ['MARCA_16','JEEP',''],['MARCA_17','MITSUBISHI',''],['MARCA_18','SUZUKI',''],['MARCA_19','DODGE',''],['MARCA_20','SUBARU',''],
      ['CAT_GASTO_1','Alquileres',''],['CAT_GASTO_2','Sueldos',''],['CAT_GASTO_3','Comisiones',''],['CAT_GASTO_4','Publicidad',''],
      ['CAT_GASTO_5','Combustible',''],['CAT_GASTO_6','Limpieza',''],['CAT_GASTO_7','Servicios',''],['CAT_GASTO_8','Gastos Bancarios',''],
      ['CAT_GASTO_9','Obra Social',''],['CAT_GASTO_10','Aportes SUSS',''],['CAT_GASTO_11','Contador',''],['CAT_GASTO_12','Mantenimiento',''],
      ['CAT_GASTO_13','Impuestos / ABL',''],['CAT_GASTO_14','Patentamiento',''],['CAT_GASTO_15','Gastos Varios',''],['CAT_GASTO_16','Reparaciones Auto',''],
      ['MDP_1','Efectivo A',''],['MDP_2','Efectivo B',''],['MDP_3','Transferencia',''],['MDP_4','Cheque',''],['MDP_5','Débito',''],['MDP_6','Mercado Pago',''],
    ]},
    {name:'VEHICULOS',color:'#16213e',headers:['ID','DOMINIO','MARCA','MODELO','AÑO','TIPO','ESTADO','PRECIO_TOMA','PRECIO_TABLA','COSTO_TOTAL','MARKUP_CONSIGNA','EX_DUENO_NOMBRE','EX_DUENO_APELLIDO','EX_DUENO_TEL','FECHA_INGRESO','OBSERVACIONES'],data:[]},
    {name:'GASTOS_VEHICULOS',color:'#0f3460',headers:['ID','ID_VEHICULO','DOMINIO','FECHA','CATEGORIA','DESCRIPCION','PROVEEDOR','IMPORTE_ARS','IMPORTE_USD','FACTURA'],data:[]},
    {name:'VENTAS',color:'#0f3460',headers:['ID','ID_VEHICULO','DOMINIO','VEHICULO','FECHA_VENTA','COMPRADOR','VENDEDOR','PRECIO_VENTA','EFECTIVO','VEHICULO_ENTREGA','FINANCIACION','TC_DIA','COSTO_TOTAL','GANANCIA','OBSERVACIONES'],data:[]},
    {name:'CAJA',color:'#533483',headers:['ID','FECHA','DETALLE','OBSERVACION','CONCEPTO','ENTIDAD','DEBE','HABER','SALDO','MEDIO_PAGO','REFERENCIA','SUB_CAJA'],data:[]},
    {name:'GASTOS_AGENCIA',color:'#533483',headers:['ID','FECHA','PERIODO','CATEGORIA','TIPO','PROVEEDOR','DESCRIPCION','NETO','IVA','TOTAL','CANAL','MEDIO_PAGO','ID_CAJA_REF','OBSERVACIONES'],data:[]},
    {name:'CHEQUES',color:'#2b2d42',headers:['ID','NUMERO','TIPO','BANCO','LIBRADOR','BENEFICIARIO','IMPORTE','FECHA_VTO','FECHA_COBRO','ESTADO','ID_CAJA_REF','OBSERVACIONES','ENTIDAD'],data:[]},
    {name:'DOLARES',color:'#2b2d42',headers:['ID','FECHA','TIPO','MONTO_USD','TC','MONTO_ARS','CONTRAPARTE','ID_CAJA_REF','OBSERVACIONES','ENTIDAD'],data:[]},
    {name:'PERSONAS',color:'#1a1a2e',headers:['ID','NOMBRE','APELLIDO','TIPO','TELEFONO','DNI','EMAIL','OBSERVACIONES'],data:[]}
  ];
  schemas.forEach(s=>{
    let sh=SS.getSheetByName(s.name);
    if(sh){sh.clearContents();sh.clearFormats();}else sh=SS.insertSheet(s.name);
    const hr=sh.getRange(1,1,1,s.headers.length);
    hr.setValues([s.headers]).setFontWeight('bold').setBackground(s.color).setFontColor('#fff').setFontSize(10);
    if(s.data.length)sh.getRange(2,1,s.data.length,s.headers.length).setValues(s.data);
    sh.setFrozenRows(1);for(let i=1;i<=s.headers.length;i++)sh.autoResizeColumn(i);
  });
  ['Sheet1','Hoja 1','Hoja1'].forEach(n=>{const s=SS.getSheetByName(n);if(s&&SS.getSheets().length>1)try{SS.deleteSheet(s);}catch(e){}});
  ui.alert('✅ AGA v2.1 listo','9 hojas creadas.\nAbrí CONFIG y reemplazá los emails.',ui.ButtonSet.OK);
}
function onOpen(){
  SpreadsheetApp.getUi().createMenu('⚙️ AGA Sistema')
    .addItem('🚀 Setup / Reset','setupSheets').addSeparator()
    .addItem('🌐 Abrir web app','abrirWebApp')
    .addItem('💱 Actualizar tipo de cambio','menuActualizarTC').addToUi();
}
function abrirWebApp(){const url=ScriptApp.getService().getUrl();SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(`<script>window.open('${url}','_blank');google.script.host.close();</script>`),'Abriendo...');}
function menuActualizarTC(){const ui=SpreadsheetApp.getUi();const b=ui.prompt('TC','Dólar blue:',ui.ButtonSet.OK_CANCEL);if(b.getSelectedButton()!==ui.Button.OK)return;const o=ui.prompt('TC','Dólar oficial:',ui.ButtonSet.OK_CANCEL);if(o.getSelectedButton()!==ui.Button.OK)return;actualizarTC(Number(b.getResponseText()),Number(o.getResponseText()));ui.alert('✅ TC actualizado.');}
