
/* ── Minimal XLSX shim (offline, no external dependency) ──
   Implements: XLSX.utils.aoa_to_sheet, XLSX.utils.book_new,
               XLSX.utils.book_append_sheet, XLSX.writeFile
   Output: real .xlsx file (OOXML) built via JSZip-free approach
   using a pre-built minimal zip binary assembler.
*/
var XLSX=(function(){
  // --- tiny zip builder (DEFLATE store only, no compression) ---
  function uint32LE(n){var b=new Uint8Array(4);b[0]=n&0xff;b[1]=(n>>8)&0xff;b[2]=(n>>16)&0xff;b[3]=(n>>24)&0xff;return b;}
  function uint16LE(n){var b=new Uint8Array(2);b[0]=n&0xff;b[1]=(n>>8)&0xff;return b;}
  function str2ab(s){var buf=new Uint8Array(s.length);for(var i=0;i<s.length;i++)buf[i]=s.charCodeAt(i)&0xff;return buf;}
  function concat(){var arrs=Array.from(arguments),total=arrs.reduce((s,a)=>s+a.length,0),out=new Uint8Array(total),pos=0;arrs.forEach(a=>{out.set(a,pos);pos+=a.length;});return out;}
  function crc32(data){
    var t=[],c,k;for(var n=0;n<256;n++){c=n;for(k=0;k<8;k++)c=c&1?(0xEDB88320^(c>>>1)):(c>>>1);t[n]=c;}
    var crc=0xFFFFFFFF;for(var i=0;i<data.length;i++)crc=(crc>>>8)^t[(crc^data[i])&0xFF];return(crc^0xFFFFFFFF)>>>0;
  }
  function makeZipEntry(name,data){
    var nameBytes=str2ab(name);
    var localHdr=concat(
      str2ab("PK\x03\x04"),uint16LE(20),uint16LE(0),uint16LE(0),uint16LE(0),uint16LE(0),
      uint32LE(crc32(data)),uint32LE(data.length),uint32LE(data.length),
      uint16LE(nameBytes.length),uint16LE(0),nameBytes,data
    );
    return localHdr;
  }
  function buildZip(files){
    var entries=[],cdEntries=[],offset=0;
    files.forEach(function(f){
      var nameBytes=str2ab(f.name),data=f.data,crc=crc32(data);
      var local=concat(str2ab("PK\x03\x04"),uint16LE(20),uint16LE(0),uint16LE(0),uint16LE(0),uint16LE(0),
        uint32LE(crc),uint32LE(data.length),uint32LE(data.length),uint16LE(nameBytes.length),uint16LE(0),nameBytes,data);
      var cd=concat(str2ab("PK\x01\x02"),uint16LE(20),uint16LE(20),uint16LE(0),uint16LE(0),uint16LE(0),uint16LE(0),
        uint32LE(crc),uint32LE(data.length),uint32LE(data.length),uint16LE(nameBytes.length),uint16LE(0),uint16LE(0),
        uint16LE(0),uint16LE(0),uint32LE(0),uint32LE(offset),nameBytes);
      entries.push(local);cdEntries.push(cd);offset+=local.length;
    });
    var cdData=concat(...cdEntries);
    var eocd=concat(str2ab("PK\x05\x06"),uint16LE(0),uint16LE(0),
      uint16LE(files.length),uint16LE(files.length),
      uint32LE(cdData.length),uint32LE(offset),uint16LE(0));
    return concat(...entries,cdData,eocd);
  }

  // --- XML helpers ---
  function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
  function colName(c){var s='';do{s=String.fromCharCode(65+(c%26))+s;c=Math.floor(c/26)-1;}while(c>=0);return s;}
  function cellRef(r,c){return colName(c)+(r+1);}

  function aoa_to_sheet(aoa){
    var ws={_aoa:aoa,_ref:null};
    var maxC=0;
    aoa.forEach(function(row){if(row.length>maxC)maxC=row.length;});
    ws._ref='A1:'+cellRef(aoa.length-1,maxC-1);
    return ws;
  }
  function book_new(){return{SheetNames:[],Sheets:{}};}
  function book_append_sheet(wb,ws,name){wb.SheetNames.push(name);wb.Sheets[name]=ws;}

  function sheetToXML(ws,name){
    var aoa=ws._aoa;
    var rows=aoa.map(function(row,ri){
      var cells=row.map(function(val,ci){
        var ref=cellRef(ri,ci);
        if(val===null||val===undefined||val==='')return'<c r="'+ref+'"><v></v></c>';
        if(typeof val==='number'||(!isNaN(parseFloat(val))&&isFinite(val)&&val!=='')){
          return'<c r="'+ref+'" t="n"><v>'+parseFloat(val)+'</v></c>';
        }
        return'<c r="'+ref+'" t="inlineStr"><is><t>'+esc(String(val))+'</t></is></c>';
      }).join('');
      return'<row r="'+(ri+1)+'">'+cells+'</row>';
    }).join('');
    return'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+
      '<sheetData>'+rows+'</sheetData></worksheet>';
  }

  function writeFile(wb,filename){
    var sheetXMLs=wb.SheetNames.map(function(n,i){
      return{name:'xl/worksheets/sheet'+(i+1)+'.xml',data:str2ab(sheetToXML(wb.Sheets[n],n))};
    });
    var sheetRels=wb.SheetNames.map(function(n,i){
      return'<Override PartName="/xl/worksheets/sheet'+(i+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    }).join('');
    var sheetRelRefs=wb.SheetNames.map(function(n,i){
      return'<Relationship Id="rId'+(i+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+(i+1)+'.xml"/>';
    }).join('');
    var wbXML='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+
      '<sheets>'+wb.SheetNames.map(function(n,i){return'<sheet name="'+esc(n)+'" sheetId="'+(i+1)+'" r:id="rId'+(i+1)+'"/>';}).join('')+'</sheets>'+
      '</workbook>';
    var ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
      '<Default Extension="xml" ContentType="application/xml"/>'+
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+
      sheetRels+
      '</Types>';
    var rootRel='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'+
      '</Relationships>';
    var wbRel='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
      sheetRelRefs+
      '</Relationships>';

    var files=[
      {name:'[Content_Types].xml',data:str2ab(ct)},
      {name:'_rels/.rels',data:str2ab(rootRel)},
      {name:'xl/workbook.xml',data:str2ab(wbXML)},
      {name:'xl/_rels/workbook.xml.rels',data:str2ab(wbRel)},
    ].concat(sheetXMLs);

    var zipData=buildZip(files);
    var blob=new Blob([zipData],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=filename;
    document.body.appendChild(a);a.click();setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(a.href);},1000);
  }

  return{
    utils:{aoa_to_sheet:aoa_to_sheet,book_new:book_new,book_append_sheet:book_append_sheet},
    writeFile:writeFile
  };
})();
