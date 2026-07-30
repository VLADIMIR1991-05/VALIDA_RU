(function(root,factory){const api=factory();if(typeof module==="object"&&module.exports)module.exports=api;else root.ValidaRU=api})(typeof self!=="undefined"?self:this,function(){
  const norm=v=>String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase();
  const key=v=>norm(v).replace(/[^A-Z0-9]/g,"");
  const number=v=>{const n=Number(String(v??"").replace(",", "."));return Number.isFinite(n)?n:null};
  const op=v=>String(v??"").replace(/\D/g,"").replace(/^0+/,"");
  const cleanHeader=v=>norm(v).toLowerCase().replace(/[^a-z0-9]+/g,"");
  const aliases={
    mueble:["codmueble","codmueblenue","codigomueble"],pieza:["tipopieza","codpieza"],cantidad:["cantreal","cantidad","cantpiezas"],
    medida1:["medida1","largo","alto"],medida2:["medida2","ancho","fondo"],stock:["codstock","codigoarticulo"],
    descripcion:["descripcio","descripcion"],orden:["op","numordent","ordenproduccion"],numero:["nomueble","numeromueble"]
  };
  function field(row,name){const wanted=aliases[name]||[];for(const [k,v] of Object.entries(row)){if(wanted.includes(cleanHeader(k)))return v}return ""}
  function detectType(headers,fileName=""){
    const h=headers.map(cleanHeader), n=norm(fileName);
    if(h.includes("codmueblenue")||h.includes("descriruta")||n.includes("HERRAJE"))return "herrajes";
    if(n.includes("LACA")||n.includes("ZICOM"))return "lacas";
    if(n.includes("MELAM"))return "melamina";
    if(h.includes("tipopieza")&&h.includes("medida1"))return "piezas";
    return "desconocido";
  }
  function parseBase(text){
    return String(text||"").split(/\r?\n/).map((raw,i)=>({line:i+1,raw,norm:norm(raw),compact:key(raw),numbers:(raw.match(/-?\d+(?:[.,]\d+)?/g)||[]).map(number)})).filter(x=>x.norm);
  }
  function equivalentCode(baseCode,excelCode){
    const a=key(baseCode),b=key(excelCode);return !!a&&!!b&&(a===b||b.startsWith(a)||a.startsWith(b));
  }
  function containsNumber(line,value,tolerance=.01){
    const n=number(value);if(n===null)return true;
    return line.numbers.some(x=>x!==null&&Math.abs(x-n)<=tolerance);
  }
  function validate(rows,baseLines,type,expectedOp){
    const issues=[];let matched=0;
    rows.forEach((row,index)=>{
      const excelOp=op(field(row,"orden"));
      const mueble=field(row,"mueble"), pieza=field(row,"pieza"), stock=field(row,"stock");
      const qty=field(row,"cantidad"),m1=field(row,"medida1"),m2=field(row,"medida2");
      if(expectedOp&&excelOp&&excelOp!==op(expectedOp)){issues.push({row:index+2,kind:"OP",detail:`OP del archivo ${excelOp} no coincide con ${op(expectedOp)}`});return}
      const candidates=baseLines.filter(l=>{
        const hasFurniture=!mueble||equivalentCode(l.compact,mueble)||l.compact.includes(key(mueble).slice(0,Math.max(5,key(mueble).indexOf("SFR"))));
        const identity=type==="herrajes"?(stock&&l.compact.includes(key(stock))):(pieza&&l.compact.includes(key(pieza)));
        return hasFurniture&&identity;
      });
      if(!candidates.length){issues.push({row:index+2,kind:"Faltante",detail:`No se encontró ${type==="herrajes"?stock:pieza} del mueble ${mueble}`});return}
      const measureMatch=candidates.some(l=>{
        if(type==="herrajes")return containsNumber(l,qty,.011);
        if(norm(pieza)==="SOPFRE")return containsNumber(l,m1,1)||containsNumber(l,number(m1)/1000,.011);
        return containsNumber(l,m1,1)&&containsNumber(l,m2,1)&&containsNumber(l,qty,.011);
      });
      if(!measureMatch){issues.push({row:index+2,kind:"Diferencia",detail:`Existe la referencia, pero cantidad o medidas no coinciden (${qty}; ${m1} × ${m2})`});return}
      matched++;
    });
    const total=rows.length,errors=issues.length,status=errors===0?"APROBADO":matched>0?"CON OBSERVACIONES":"RECHAZADO";
    return {total,matched,errors,missing:issues.filter(x=>x.kind==="Faltante").length,differences:issues.filter(x=>x.kind!=="Faltante").length,status,issues};
  }
  return {norm,key,op,field,detectType,parseBase,equivalentCode,validate};
});
