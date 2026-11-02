/**
 * INFORMES TIERRA QUÍN → Google Sheets + correo
 * - informeRango(): crea el informe de las fechas de abajo (INF_DESDE / INF_HASTA) y lo envía por correo.
 * - informeMesAnterior(): crea y envía el informe del mes que terminó. Se ejecuta solo el día 1 de cada mes.
 * - instalarInformes(): ejecútalo UNA vez para programar el informe automático de cada mes.
 * No cambia nada en el sistema: solo lee la información y crea la hoja.
 */
const INF_DESDE = '2026-09-23';
const INF_HASTA = '2026-10-31';
const INF_CORREOS = [];            // Ej. ['paloma@gmail.com','juliana@gmail.com']. Si queda vacío, se envía a la cuenta de Tierra Quín.
const INF_CARPETA = 'Informes Tierra Quin';

function informeRango(){ generarInforme_(INF_DESDE, INF_HASTA, true); }
function informeMesAnterior(){
  const tz='America/Bogota', h=new Date();
  const a=new Date(h.getFullYear(), h.getMonth()-1, 1), b=new Date(h.getFullYear(), h.getMonth(), 0);
  generarInforme_(Utilities.formatDate(a,tz,'yyyy-MM-dd'), Utilities.formatDate(b,tz,'yyyy-MM-dd'), true);
}
function instalarInformes(){
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='informeMesAnterior').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('informeMesAnterior').timeBased().onMonthDay(1).atHour(7).inTimezone('America/Bogota').create();
  console.log('Listo: el informe del mes se crea y se envía solo cada día 1 a las 7 a. m.');
}

function generarInforme_(desde, hasta, enviar){
  const D = leerTodo_();
  const inf = construirInforme(D, desde, hasta);
  const titulo = 'Tierra Quín · Informe ' + desde + ' a ' + hasta;
  const ss = SpreadsheetApp.create(titulo);
  const tablas = tablasInforme(inf);
  tablas.forEach(([nombre, filas], k)=>{
    const sh = k===0 ? ss.getSheets()[0].setName(nombre) : ss.insertSheet(nombre);
    if(!filas.length) return;
    const ancho = Math.max.apply(null, filas.map(f=>f.length));
    const datos = filas.map(f=>{ const x=f.slice(); while(x.length<ancho) x.push(''); return x.map(v=>v==null?'':v); });
    sh.getRange(1,1,datos.length,ancho).setValues(datos);
    if(nombre==='Resumen'){ sh.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#D6336C'); sh.getRange(1,2,datos.length,1).setNumberFormat('"$"#,##0');
      ['Mercados entregados','Órdenes entregadas','Órdenes canceladas','Órdenes que no cuentan en Finanzas','Periodo'].forEach(t=>{ datos.forEach((f,i)=>{ if(f[0]===t) sh.getRange(i+1,2).setNumberFormat('0'); }); });
      datos.forEach((f,i)=>{ if(f[0]==='GANANCIA') sh.getRange(i+1,1,1,2).setFontWeight('bold').setBackground('#FDE4EE'); });
    } else {
      sh.getRange(1,1,1,ancho).setFontWeight('bold').setBackground('#E6EEFD').setFontColor('#1D5BD6'); sh.setFrozenRows(1);
      if(datos.length>1) for(let c=0;c<ancho;c++){ const v=datos.slice(1).find(f=>typeof f[c]==='number'); if(v!==undefined && !/Cantidad|Mercados|Órdenes/.test(datos[0][c])) sh.getRange(2,c+1,datos.length-1,1).setNumberFormat('"$"#,##0'); }
    }
    sh.autoResizeColumns(1, ancho);
  });
  SpreadsheetApp.flush();
  const archivo = DriveApp.getFileById(ss.getId());
  const it = DriveApp.getFoldersByName(INF_CARPETA); const carpeta = it.hasNext()?it.next():DriveApp.createFolder(INF_CARPETA);
  archivo.moveTo(carpeta);
  console.log('Hoja creada: '+ss.getUrl());
  const r=inf.resumen; const c=n=>'$'+Math.round(n).toLocaleString('es-CO');
  console.log('Ventas '+c(r.ventas)+' · Costo '+c(r.costo)+' · Gastos '+c(r.gastos)+' · Ajustes '+c(r.ajustes)+' · Ganancia '+c(r.ganancia)+' · Cobrado '+c(r.cobrado)+' · Inversión '+c(r.inversion));
  if(!enviar) return ss.getUrl();
  const para = (INF_CORREOS.length?INF_CORREOS:[Session.getEffectiveUser().getEmail()]).join(',');
  const xlsx = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/'+ss.getId()+'/export?format=xlsx',{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}}).getBlob().setName(titulo.replace(/[·]/g,'-')+'.xlsx');
  const html = '<div style="font-family:Arial,sans-serif;color:#1F2340"><h2 style="color:#D6336C">Tierra Quín · Informe '+desde+' a '+hasta+'</h2>'
    +'<table cellpadding="6" style="border-collapse:collapse">'
    +[['Mercados entregados',r.mercados,1],['Ventas (sin domicilios)',r.ventas],['Costo del producto',r.costo],['Otros gastos',r.gastos],['Ajustes',r.ajustes],['<b>Ganancia</b>',r.ganancia],['Domicilios cobrados',r.domicilios],['Propinas',r.propinas],['Cobrado',r.cobrado],['Por cobrar',r.porCobrar],['Inversión en compras',r.inversion]]
      .map(x=>'<tr><td style="border-bottom:1px solid #eee">'+x[0]+'</td><td style="border-bottom:1px solid #eee;text-align:right"><b>'+(x[2]?x[1]:c(x[1]))+'</b></td></tr>').join('')
    +'</table><p><a href="'+ss.getUrl()+'" style="background:#D6336C;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">Abrir el informe en Google Sheets</a></p>'
    +'<p style="color:#666;font-size:12px">También va adjunto en Excel. La hoja queda guardada en Google Drive, carpeta "'+INF_CARPETA+'".</p><p>🩵🩷</p></div>';
  MailApp.sendEmail({to:para, subject:titulo, htmlBody:html, attachments:[xlsx]});
  console.log('Correo enviado a: '+para);
  return ss.getUrl();
}

function leerTodo_(){
  const token=ScriptApp.getOAuthToken(); const D={};
  ['cli','prod','ordenes','combos','gastos','config','compras','combosDia'].forEach(c=>{
    D[c]={}; let pageToken='';
    do{
      let url='https://firestore.googleapis.com/v1/projects/'+PROYECTO+'/databases/(default)/documents/'+c+'?pageSize=300';
      if(pageToken) url+='&pageToken='+encodeURIComponent(pageToken);
      const r=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+token,'x-goog-user-project':PROYECTO},muteHttpExceptions:true});
      if(r.getResponseCode()!==200) throw new Error('No se pudo leer "'+c+'": '+r.getContentText().slice(0,300));
      const j=JSON.parse(r.getContentText()||'{}');
      (j.documents||[]).forEach(doc=>{ D[c][decodeURIComponent(doc.name.split('/').pop())]=campos(doc.fields||{}); });
      pageToken=j.nextPageToken||'';
    } while(pageToken);
  });
  return D;
}

/* ===== MOTOR DEL INFORME (igual en el sistema y en Google) =====
   D = {cli, prod, ordenes, combosDia, compras, gastos, config}; desde/hasta = 'AAAA-MM-DD' */
function construirInforme(D, desde, hasta){
  var prods=D.prod||{}, combosDia=D.combosDia||{}, compras=D.compras||{}, gastos=D.gastos||{}, clientes=D.cli||{};
  var cfg=((D.config||{}).negocio)||{}; var adics=Array.isArray(cfg.adicionales)?cfg.adicionales:[];
  var cmp=function(a,b){ return a<b?-1:a>b?1:0; };
  var cant=function(o){ return Math.max(1,Math.round(+(o&&o.cantidad)||1)); };
  var hist=function(p){ return (p.historial||[]).slice().sort(function(a,b){ return cmp(a.f,b.f); }); };
  var precioMercado=function(pid,f,ref){ var p=prods[pid]; if(!p) return +ref||0; var h=hist(p); var e=null; h.forEach(function(x){ if(x.f<=f) e=x; }); if(e) return +e.p; return h.length?+h[0].p:(+p.precio||0); };
  var precio=function(pid,f,ref){ var cp=((compras[f]||{}).items||{})[pid]; if(cp && cp.pagado!=null && cp.pagado!=='') return +cp.pagado; return precioMercado(pid,f,ref); };
  var itemsDe=function(o){ if(o.personalizado && Array.isArray(o.items) && o.items.length) return o.items; var d=combosDia[o.fecha+'_'+((o.combo||{}).id)]; return (d && Array.isArray(d.items))?d.items:[]; };
  var costoUnit=function(o){ return itemsDe(o).reduce(function(t,i){ return t+(+i.cant||0)*precio(i.productoId,o.fecha,i.precioRef); },0); };
  var adicTotal=function(o){ return (o.adicionales||[]).reduce(function(t,a){ return t+(+a.cant||0)*(+a.precio||0); },0); };
  var costoAdic=function(a,f){ var cp=((compras[f]||{}).items||{})['ad:'+a.id]; if(cp && cp.pagado!=null && cp.pagado!=='') return +cp.pagado;
    var cur=adics.filter(function(x){ return x.id===a.id; })[0];
    if(cur && Array.isArray(cur.hist) && cur.hist.length){ var e=null; cur.hist.slice().sort(function(x,y){ return cmp(x.f,y.f); }).forEach(function(x){ if(x.f<=f) e=x; }); if(e) return +e.c; }
    if(+a.costo) return +a.costo; return cur?(+cur.costo||0):0; };
  var adicCosto=function(o){ return (o.adicionales||[]).reduce(function(t,a){ return t+(+a.cant||0)*costoAdic(a,o.fecha); },0); };
  var extras=function(o){ return o.extrasMercado||[]; };
  var ventaExtras=function(o){ return extras(o).reduce(function(t,x){ return t+(+x.cant||0)*(+x.precio||0); },0); };
  var costoExtras=function(o){ return extras(o).reduce(function(t,x){ return t+(+x.cant||0)*precio(x.productoId,o.fecha,x.precioRef); },0); };
  var ventaCombo=function(o){ return (+((o.combo||{}).precio)||0)*cant(o); };
  var costoTotal=function(o){ return costoUnit(o)*cant(o)+adicCosto(o)+costoExtras(o); };
  var domExt=function(o){ return o.domTipo==='propio'?0:(+o.domicilio||0); };   // domiciliario externo: se le paga
  var domProp=function(o){ return o.domTipo==='propio'?(+o.domicilio||0):0; };   // lo llevan ellas: es ganancia
  var ventaNeta=function(o){ return (+o.total||0)-domExt(o); };
  var aCobrar=function(o){ return Math.max(0,(+o.total||0)-(+o.saldoUsado||0)); };
  var nombre=function(o){ return [o.nombre,o.apellido].filter(function(x){ return x && x!=='.'; }).join(' ') || (o.nombre||''); };
  var num=function(n){ return n?'TQ-'+String(n).padStart(4,'0'):''; };
  var EST={pendiente:'Pendiente de entrega',preparacion:'En preparación',camino:'En camino',demorado:'Demorado',reprogramado:'Reprogramado',entregado:'Entregado',cancelado:'Cancelado'};
  var todas=Object.keys(D.ordenes||{}).map(function(id){ return [id,D.ordenes[id]]; }).filter(function(x){ return x[1].fecha>=desde && x[1].fecha<=hasta; })
    .sort(function(a,b){ return cmp(a[1].fecha,b[1].fecha)||((+a[1].numero||0)-(+b[1].numero||0)); });
  var ps=todas.map(function(x){ return x[1]; }).filter(function(o){ return o.estado!=='cancelado' && !o.excluir; });
  var ent=ps.filter(function(o){ return o.estado==='entregado'; });
  var sum=function(l,fn){ return l.reduce(function(t,o){ return t+fn(o); },0); };
  // Compras de un día (igual que Inversión)
  var ordenesDia=function(f){ return Object.keys(D.ordenes||{}).map(function(id){ return D.ordenes[id]; }).filter(function(o){ return o.fecha===f && o.estado!=='cancelado'; }); };
  var comprasDia=function(f){
    var agg={}, ad={};
    ordenesDia(f).forEach(function(o){ var k=cant(o);
      itemsDe(o).forEach(function(i){ var a=agg[i.productoId]||(agg[i.productoId]={n:i.nombre,u:i.unidad,c:0,ref:i.precioRef}); a.c+=(+i.cant||0)*k; });
      extras(o).forEach(function(i){ var a=agg[i.productoId]||(agg[i.productoId]={n:i.nombre,u:i.unidad,c:0,ref:i.precioRef}); a.c+=(+i.cant||0); });
      (o.adicionales||[]).forEach(function(x){ if(x.cant>0){ var a=ad[x.id]||(ad[x.id]={n:x.nombre,c:0}); a.c+=+x.cant||0; } }); });
    var doc=compras[f]||{}; var st=doc.items||{}; var filas=[];
    var ids={}; Object.keys(agg).forEach(function(k){ ids[k]=1; }); Object.keys(st).forEach(function(k){ if(st[k].agregado && k.indexOf('ad:')!==0) ids[k]=1; });
    Object.keys(ids).forEach(function(pid){ var a=agg[pid]; var p=prods[pid]; var x=st[pid]||{}; var nec=a?a.c:0; var q=x.cant!=null?+x.cant:nec;
      var pag=(x.pagado!=null && x.pagado!=='')?+x.pagado:null; var pr=pag!=null?pag:precioMercado(pid,f,a?a.ref:0);
      filas.push({fecha:f,producto:p?p.nombre:(a?a.n:'Producto'),unidad:p?p.unidad:(a?a.u:''),cantidad:q,precio:pr,real:pag!=null,total:q*pr}); });
    adics.forEach(function(a){ var key='ad:'+a.id; var need=ad[a.id]?ad[a.id].c:0; var x=st[key]||{}; if(!need && !x.agregado) return;
      var q=x.cant!=null?+x.cant:need; var pag=(x.pagado!=null && x.pagado!=='')?+x.pagado:null; var pr=pag!=null?pag:(+a.costo||0);
      filas.push({fecha:f,producto:a.nombre+(a.detalle?' ('+a.detalle+')':''),unidad:'und',cantidad:q,precio:pr,real:pag!=null,total:q*pr}); });
    (doc.extras||[]).forEach(function(x){ filas.push({fecha:f,producto:x.nombre,unidad:x.unidad||'',cantidad:+x.cant||0,precio:+x.precio||0,real:true,total:(+x.cant||0)*(+x.precio||0)}); });
    filas.sort(function(a,b){ return a.producto.localeCompare(b.producto,'es'); });
    return filas;
  };
  var dias=[]; var d0=new Date(desde+'T12:00:00'), d1=new Date(hasta+'T12:00:00');
  for(var d=new Date(d0); d<=d1; d.setDate(d.getDate()+1)){ dias.push(d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')); }
  var compRows=[]; var invDia={};
  dias.forEach(function(f){ var hay=ps.some(function(o){ return o.fecha===f; }) || compras[f]; if(!hay){ invDia[f]=0; return; } var fl=comprasDia(f); invDia[f]=sum(fl,function(x){ return x.total; }); compRows=compRows.concat(fl); });
  var todosMov=Object.keys(gastos).map(function(id){ return gastos[id]; }).filter(function(g){ return g.fecha>=desde && g.fecha<=hasta; }).sort(function(a,b){ return cmp(a.fecha,b.fecha); });
  var socios=todosMov.filter(function(g){ return g.socio==='aporte'||g.socio==='reembolso'; });   // plata de las socias: no es gasto
  var gas=todosMov.filter(function(g){ return !(g.socio==='aporte'||g.socio==='reembolso'); });
  var socTot=function(t){ return sum(socios.filter(function(g){ return g.socio===t; }),function(g){ return +g.monto||0; }); };
  var socHist=function(t){ return Object.keys(gastos).map(function(id){ return gastos[id]; }).filter(function(g){ return g.socio===t && g.fecha<=hasta; }).reduce(function(a,g){ return a+(+g.monto||0); },0); };
  var gT=sum(gas,function(g){ return +g.monto||0; });
  var esAjuste=function(g){ return String(g.categoria||'').indexOf('Ajuste')===0; };
  var metodos=['Efectivo','Nequi','Daviplata','Bancolombia'].map(function(m){ var l=ent.filter(function(o){ return o.pagado && (o.metodo||'Efectivo')===m; }); return {metodo:m,ordenes:l.length,valor:sum(l,function(o){ return aCobrar(o)+(+o.propina||0); })}; });
  var abonos=0; Object.keys(clientes).forEach(function(cid){ (clientes[cid].movs||[]).forEach(function(mv){ if(mv.tipo==='abono' && mv.f>=desde && mv.f<=hasta) abonos+=+mv.m||0; }); });
  var ventas=sum(ent,ventaNeta), costo=sum(ent,costoTotal);
  var res={
    desde:desde, hasta:hasta,
    mercados:sum(ent,cant), ordenes:ent.length,
    ventas:ventas, ventasCombos:sum(ent,ventaCombo), ventasAdicionales:sum(ent,adicTotal), ventasExtras:sum(ent,ventaExtras),
    costo:costo, gastos:sum(gas.filter(function(g){ return !esAjuste(g); }),function(g){ return +g.monto||0; }), ajustes:sum(gas.filter(esAjuste),function(g){ return +g.monto||0; }),
    ganancia:ventas-costo-gT,
    domicilios:sum(ent,domExt), domiciliosPropios:sum(ent,domProp), propinas:sum(ent,function(o){ return +o.propina||0; }),
    cobrado:sum(ent.filter(function(o){ return o.pagado; }),aCobrar), porCobrar:sum(ent.filter(function(o){ return !o.pagado; }),aCobrar),
    inversion:dias.reduce(function(t,f){ return t+invDia[f]; },0),
    saldoFavorRecibido:abonos, saldoUsado:sum(ent,function(o){ return +o.saldoUsado||0; }),
    canceladas:todas.filter(function(x){ return x[1].estado==='cancelado'; }).length,
    noCuentan:todas.filter(function(x){ return x[1].excluir; }).length,
    porEntregar:sum(ps.filter(function(o){ return o.estado!=='entregado'; }),function(o){ return +o.total||0; }),
    aportes:socTot('aporte'), reembolsos:socTot('reembolso'), deudaSocias:socHist('aporte')-socHist('reembolso')
  };
  var ventasRows=todas.map(function(x){ var o=x[1]; var cuenta=o.estado==='entregado' && !o.excluir;
    return {fecha:o.fecha,orden:num(o.numero),cliente:nombre(o),telefono:[o.telefono,o.telefono2].filter(Boolean).join(' / '),direccion:o.direccion||'',
      combo:((o.combo||{}).nombre)||'',cantidad:cant(o),valorCombos:ventaCombo(o),
      adicionales:(o.adicionales||[]).filter(function(a){ return a.cant>0; }).map(function(a){ return a.cant+' x '+a.nombre; }).join(', '),valorAdicionales:adicTotal(o),
      extras:extras(o).map(function(e){ return e.cant+' '+(e.unidad||'')+' '+e.nombre; }).join(', '),valorExtras:ventaExtras(o),
      domicilio:+o.domicilio||0,domQuien:(+o.domicilio?(o.domTipo==='propio'?'Nosotras (ganancia)':'Domiciliario externo'):''),saldoUsado:+o.saldoUsado||0,total:+o.total||0,aCobrar:aCobrar(o),propina:+o.propina||0,
      estado:EST[o.estado]||o.estado||'',pagado:o.pagado?'Sí':'No',metodo:o.pagado?(o.metodo||'Efectivo'):'',
      costo:costoTotal(o),ganancia:ventaNeta(o)-costoTotal(o),cuenta:cuenta?'Sí':(o.excluir?'No (excluida)':'No (sin entregar)')}; });
  var diario=dias.map(function(f){ var l=ent.filter(function(o){ return o.fecha===f; }); var gd=gas.filter(function(g){ return g.fecha===f && !esAjuste(g); });
    var ing=sum(l,function(o){ return +o.total||0; }), dom=sum(l,domExt), comp=invDia[f]||0, gst=sum(gd,function(g){ return +g.monto||0; });
    return {fecha:f,mercados:sum(l,cant),ingresos:ing,domicilios:dom,compras:comp,gastos:gst,egresos:comp+gst+dom,neto:ing-comp-gst-dom}; }).filter(function(r){ return r.ingresos||r.compras||r.gastos; });
  var gastosRows=gas.map(function(g){ return {fecha:g.fecha,tipo:g.categoria||'',descripcion:g.descripcion||'',valor:+g.monto||0}; });
  var sociosRows=socios.map(function(g){ return {fecha:g.fecha,tipo:g.socio==='aporte'?'Aporte (puso plata)':'Reembolso (se le devolvió)',socia:g.socia||'',descripcion:g.descripcion||'',valor:+g.monto||0}; });
  var saldos=Object.keys(clientes).map(function(cid){ return clientes[cid]; }).filter(function(c){ return +c.saldo>0; }).map(function(c){ return {cliente:nombre(c),saldo:+c.saldo}; });
  return {resumen:res, ventas:ventasRows, diario:diario, compras:compRows, gastos:gastosRows, socios:sociosRows, metodos:metodos, saldos:saldos};
}
/* Tablas listas para hoja de cálculo */
function tablasInforme(inf){
  var r=inf.resumen;
  var resumen=[['Informe Tierra Quín',''],['Periodo',r.desde+' a '+r.hasta],['',''],
    ['Mercados entregados',r.mercados],['Órdenes entregadas',r.ordenes],
    ['Ventas (sin domicilios externos)',r.ventas],['   Combos',r.ventasCombos],['   Adicionales',r.ventasAdicionales],['   Extras del mercado',r.ventasExtras],
    ['Costo del producto vendido',r.costo],['Otros gastos',r.gastos],['Ajustes para cuadrar',r.ajustes],['GANANCIA',r.ganancia],['',''],
    ['   Domicilios propios (incluidos en ventas)',r.domiciliosPropios],['Domicilios externos (para el domiciliario)',r.domicilios],['Propinas',r.propinas],['Saldo a favor recibido de clientes',r.saldoFavorRecibido],['Saldo a favor usado en órdenes',r.saldoUsado],['',''],
    ['Cobrado',r.cobrado],['Por cobrar',r.porCobrar]].concat(inf.metodos.map(function(m){ return ['   Cobrado en '+m.metodo,m.valor]; })).concat([
    ['',''],['Inversión en compras del mercado',r.inversion],['Órdenes canceladas',r.canceladas],['Órdenes que no cuentan en Finanzas',r.noCuentan],
    ['',''],['Aportes de las socias en el periodo',r.aportes],['Reembolsos a las socias en el periodo',r.reembolsos],['Lo que el negocio les debe a las socias (al final)',r.deudaSocias]]);
  var V=inf.ventas; var ventas=[['Fecha','Orden','Cliente','Teléfono','Dirección','Combo','Cantidad','Valor combos','Adicionales','Valor adicionales','Extras','Valor extras','Domicilio','¿Quién lleva el domicilio?','Saldo a favor usado','Total','Total a cobrar','Propina','Estado','Pagado','Método de pago','Costo','Ganancia','Cuenta en Finanzas']]
    .concat(V.map(function(x){ return [x.fecha,x.orden,x.cliente,x.telefono,x.direccion,x.combo,x.cantidad,x.valorCombos,x.adicionales,x.valorAdicionales,x.extras,x.valorExtras,x.domicilio,x.domQuien,x.saldoUsado,x.total,x.aCobrar,x.propina,x.estado,x.pagado,x.metodo,x.costo,x.ganancia,x.cuenta]; }));
  var diario=[['Fecha','Mercados','Ingresos del día','Domicilios externos (se pagan al domiciliario)','Compras del mercado','Otros gastos','Total egresos','Neto del día']]
    .concat(inf.diario.map(function(x){ return [x.fecha,x.mercados,x.ingresos,x.domicilios,x.compras,x.gastos,x.egresos,x.neto]; }));
  var compras=[['Fecha','Producto','Unidad','Cantidad','Precio','¿Precio real pagado?','Total']].concat(inf.compras.map(function(x){ return [x.fecha,x.producto,x.unidad,x.cantidad,x.precio,x.real?'Sí':'Estimado (Mercado)',x.total]; }));
  var gastos=[['Fecha','Tipo','Descripción','Valor']].concat(inf.gastos.map(function(x){ return [x.fecha,x.tipo,x.descripcion,x.valor]; }));
  var socias=[['Fecha','Movimiento','Socia','Descripción','Valor']].concat((inf.socios||[]).map(function(x){ return [x.fecha,x.tipo,x.socia,x.descripcion,x.valor]; }));
  var metodos=[['Método de pago','Órdenes','Valor cobrado']].concat(inf.metodos.map(function(m){ return [m.metodo,m.ordenes,m.valor]; }));
  var saldos=[['Cliente','Saldo a favor actual']].concat(inf.saldos.map(function(x){ return [x.cliente,x.saldo]; }));
  return [['Resumen',resumen],['Ventas',ventas],['Libro diario',diario],['Compras',compras],['Gastos',gastos],['Aportes de socias',socias],['Medios de pago',metodos],['Saldos a favor',saldos]];
}
