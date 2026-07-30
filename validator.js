(function(root,factory){const api=factory();if(typeof module==="object"&&module.exports)module.exports=api;else root.ValidaRU=api})(typeof self!=="undefined"?self:this,function(){
  const norm=v=>String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toUpperCase();
  const key=v=>norm(v).replace(/[^A-Z0-9]/g,"");
  const number=v=>{const n=Number(String(v??"").replace(",", "."));return Number.isFinite(n)?n:null};
  const op=v=>String(v??"").replace(/\D/g,"").replace(/^0+/,"");
  const stockKey=v=>key(v).replace(/[A-Z]+$/,"").replace(/^0+/,"");
  const pieceKey=v=>key(String(v??"").split(",")[0]);
  const furnitureKey=v=>key(v).replace(/(?:SFR|SB|FLBX|REJ)+$/,"");
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
    return String(text||"").split(/\r?\n/).map((raw,i)=>{
      const fields=[];let value="",quoted=false;
      const delimiter=raw.includes("|")?"|":";";
      for(let p=0;p<raw.length;p++){const c=raw[p];if(c==='"')quoted=!quoted;else if(c===delimiter&&!quoted){fields.push(value);value=""}else value+=c}
      fields.push(value);
      return {line:i+1,raw,norm:norm(raw),compact:key(raw),numbers:(raw.match(/-?\d+(?:[.,]\d+)?/g)||[]).map(number),
        fields,numero:fields[0]||"",mueble:fields[1]||"",description:fields[2]||"",piece:pieceKey(fields[2]),qty:number(fields[3]),m1:number(fields[4]),m2:number(fields[5]),stock:fields[8]||""};
    }).filter(x=>x.norm);
  }
  function equivalentCode(baseCode,excelCode){
    const a=key(baseCode),b=key(excelCode);return !!a&&!!b&&(a===b||b.startsWith(a)||a.startsWith(b));
  }
  function containsNumber(line,value,tolerance=.01){
    const n=number(value);if(n===null)return true;
    return line.numbers.some(x=>x!==null&&Math.abs(x-n)<=tolerance);
  }
  function comparison(row,line,type,index,note="Coincidencia"){
    const numero=field(row,"numero"),mueble=field(row,"mueble"),pieza=field(row,"pieza"),stock=field(row,"stock"),qty=field(row,"cantidad"),m1=field(row,"medida1"),m2=field(row,"medida2");
    return {row:index+2,kind:note,numero:String(numero||line?.numero||"—"),mueble:String(mueble||"—"),groupKey:`${numero||line?.numero||"—"}|${mueble||"—"}`,reference:type==="herrajes"?String(stock):String(pieza),xls:type==="herrajes"?`Cant. ${qty}`:`Cant. ${qty} · ${m1} × ${m2}`,txt:line?.raw||"Sin línea directa",txtLine:line?.line||null,detail:note};
  }
  function validate(rows,baseLines,type,expectedOp){
    const issues=[],comparisons=[];let matched=0,notComparable=0;
    const byStock=new Map(),byPiece=new Map();
    baseLines.forEach(line=>{
      const sk=stockKey(line.stock);if(sk){if(!byStock.has(sk))byStock.set(sk,[]);byStock.get(sk).push(line)}
      if(line.piece){if(!byPiece.has(line.piece))byPiece.set(line.piece,[]);byPiece.get(line.piece).push(line)}
    });
    rows.forEach((row,index)=>{
      const excelOp=op(field(row,"orden"));
      const mueble=field(row,"mueble"), pieza=field(row,"pieza"), stock=field(row,"stock");
      const qty=field(row,"cantidad"),m1=field(row,"medida1"),m2=field(row,"medida2");
      if(expectedOp&&excelOp&&excelOp!==op(expectedOp)){issues.push({row:index+2,kind:"OP",detail:`La OP ${excelOp} del XLS no coincide con la OP base ${op(expectedOp)}`,reference:type==="herrajes"?String(stock):String(pieza),xls:`Mueble ${mueble||"—"} · OP ${excelOp}`,txt:`TXT base / filtro: OP ${op(expectedOp)}`});return}
      const identityPool=type==="herrajes"?(byStock.get(stockKey(stock))||[]):(byPiece.get(pieceKey(pieza))||[]);
      const candidates=identityPool.filter(l=>!mueble||equivalentCode(furnitureKey(l.mueble),furnitureKey(mueble)));
      if(!candidates.length){
        if(type==="herrajes"){notComparable++;comparisons.push(comparison(row,null,type,index,"Sin equivalencia directa"));return}
        issues.push({...comparison(row,null,type,index,"Faltante"),detail:`No se encontró ${pieza} del mueble ${mueble}`,txt:`Sin registro equivalente en el TXT para mueble ${mueble||"—"} y pieza ${pieza||"—"}`});return
      }
      const measureMatch=candidates.some(l=>{
        if(type==="herrajes"){
          const q=number(qty);
          if(q===null||q===0)return true;
          return candidates.some(x=>Math.abs((x.qty??Infinity)-q)<=.011||Math.abs((((x.qty??0)*(x.m1??0))/1000)-q)<=.011);
        }
        if(norm(pieza)==="SOPFRE")return containsNumber(l,m1,1)||containsNumber(l,number(m1)/1000,.011);
        return containsNumber(l,m1,1)&&containsNumber(l,m2,1)&&containsNumber(l,qty,.011);
      });
      if(!measureMatch){
        if(type==="herrajes"){matched++;comparisons.push(comparison(row,candidates[0],type,index,"Referencia válida"));return}
        issues.push({...comparison(row,candidates[0],type,index,"Diferencia"),detail:`Existe la referencia, pero cantidad o medidas no coinciden`});return
      }
      matched++;
      comparisons.push(comparison(row,candidates.find(l=>{
        if(type==="herrajes")return true;
        return containsNumber(l,m1,1)&&containsNumber(l,m2,1);
      })||candidates[0],type,index,norm(pieza)==="SOPFRE"?"Regla SOPFRE":"Coincide"));
    });
    const total=rows.length,errors=issues.length,status=errors===0?"APROBADO":matched>0?"CON OBSERVACIONES":"RECHAZADO";
    return {total,matched,notComparable,errors,missing:issues.filter(x=>x.kind==="Faltante").length,differences:issues.filter(x=>x.kind!=="Faltante").length,status,issues,comparisons};
  }
  return {norm,key,op,stockKey,pieceKey,furnitureKey,field,detectType,parseBase,equivalentCode,validate};
});
