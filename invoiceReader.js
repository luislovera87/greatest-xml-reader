window.onload = function(){
  document.getElementById('exportBtn').setAttribute('disabled','disabled')

};

//XML import

var fileChooser = document.getElementById("fileChooser");
var reader;
var tbody = document.getElementById('table').children[1];
var files;
var tr;
var td;

function errorHandler(evt) {
    switch(evt.target.error.code) {
      case evt.target.error.NOT_FOUND_ERR:
        alert('File Not Found!');
        break;
      case evt.target.error.NOT_READABLE_ERR:
        alert('File is not readable');
        break;
      case evt.target.error.ABORT_ERR:
        break; // noop
      default:
        alert('An error occurred reading this file.');
    };
  }

function handleFileSelection() {
  clearTable();
    files = fileChooser.files;
    for(var i=0; i <= files.length - 1; i++){

        reader = new FileReader();
        reader.onerror = errorHandler;  
        reader.onloadend = function(event) {          
          var text = event.target.result;

          var parser = new DOMParser();
          var xmlDom = parser.parseFromString(text, "text/xml");                                                                

          var comprobante = xmlDom.getElementsByTagName('cfdi:Comprobante')[0];     
          var folio = comprobante.getAttribute('Folio');
          var moneda = comprobante.getAttribute('Moneda');
          var formaPago = comprobante.getAttribute('FormaPago');
          var metodoPago = comprobante.getAttribute('MetodoPago');

          var emisor = xmlDom.getElementsByTagName('cfdi:Emisor')[0];
          var nombreEmisor = emisor.getAttribute('Nombre');

          var receptor = xmlDom.getElementsByTagName('cfdi:Receptor')[0];
          var nombreReceptor = receptor.getAttribute('Nombre');
          var usoCFDI = receptor.getAttribute('UsoCFDI');

          var conceptos = xmlDom.getElementsByTagName('cfdi:Conceptos')[0];            
          var conceptosLen = conceptos.childNodes.length;

          var conceptosArr = [];          
          for(var i = 0; i <= conceptosLen - 1; i++){
            if(conceptos.childNodes[i].nodeType !== Node.TEXT_NODE){
              
              conceptosArr.push({
                Cantidad: conceptos.childNodes[i].getAttribute('Cantidad'),
                Descripcion: conceptos.childNodes[i].getAttribute('Descripcion'),
                Importe: conceptos.childNodes[i].getAttribute('Importe')
              });

            }                        
          }
          
          var timbreFiscalDigital = xmlDom.getElementsByTagName('tfd:TimbreFiscalDigital')[0];
          var fechaTimbrado = timbreFiscalDigital.getAttribute('FechaTimbrado').split("T")[0];

        var row = [];
        row.push({
          Vendor: nombreEmisor,
          Conceptos: conceptosArr,
          FechaTimbrado: fechaTimbrado,
          Moneda: moneda,
          Folio: folio,
          FormaPago: formaPago,
          MetodoPago: metodoPago,
          UsoCFDI: usoCFDI,
          Receptor: nombreReceptor          
        });

        var rowOb = row[0];

        for(var j=0; j <= conceptosArr.length - 1; j++){
           tr = tbody.insertRow();
            
            // Vendor
            td = tr.insertCell();  
            td.setAttribute('class','vendor');          
            td.innerHTML = rowOb.Vendor;

            // Concepto - Descripcion
            td = tr.insertCell();
            td.setAttribute('class','descripcion');
            td.innerHTML = rowOb.Conceptos[j].Descripcion

            // Fecha de Timbrado
            td = tr.insertCell();
            td.setAttribute('class','fechaTimbrado');
            td.innerHTML = rowOb.FechaTimbrado;

            // Concepto - Importe
            td = tr.insertCell();
            td.innerHTML = rowOb.Conceptos[j].Importe
            
            // Moneda
            td = tr.insertCell();            
            td.innerHTML = rowOb.Moneda;

            // Concepto - Cantidad 
            td = tr.insertCell();
            td.innerHTML = parseInt(rowOb.Conceptos[j].Cantidad)

            // Folio
            td = tr.insertCell();            
            td.innerHTML = rowOb.Folio;

            // Forma de Pago
            td = tr.insertCell();            
            td.innerHTML = rowOb.FormaPago;

            // Metodo de Pago
            td = tr.insertCell();            
            td.innerHTML = rowOb.MetodoPago;

            // Uso de CFDI
            td = tr.insertCell();            
            td.innerHTML = rowOb.UsoCFDI;

            // Receptor
            td = tr.insertCell();            
            td.innerHTML = rowOb.Receptor;
        }       
              
    }
        reader.readAsText(files[i]);        
    }     
    document.getElementById('exportBtn').removeAttribute('disabled');
}

fileChooser.addEventListener('change', handleFileSelection, false);

//Excel export

var tableToExcel = (function () {
        var uri = 'data:application/vnd.ms-excel;base64,'
        , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">'+
      '<head><!--[if gte mso 9]>'+
        '<xml>'+
          '<x:ExcelWorkbook>'+
            '<x:ExcelWorksheets>'+
              '<x:ExcelWorksheet><x:Name>{worksheet}</x:Name>'+
              '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>'+
              '</x:ExcelWorksheet>'+
            '</x:ExcelWorksheets>'+
          '</x:ExcelWorkbook></xml><![endif]-->'+
        '<meta http-equiv="content-type" content="text/plain; charset=UTF-8"/>'+
      '</head>'+
      '<body>'+
        '<table>'+
          '{table}'+
        '</table>'+
      '</body>'+
    '</html>'
        , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
        , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
        return function (table, name, filename) {
            if (!table.nodeType) table = document.getElementById(table)
            var ctx = { worksheet: name || 'Worksheet', table: table.innerHTML }

            document.getElementById("dlink").href = uri + base64(format(template, ctx));
            document.getElementById("dlink").download = filename;
            document.getElementById("dlink").click();

        }
    })()

function clearTable(){
  var tableRows = document.querySelectorAll('table tbody tr td');
  
  tableRows.forEach(function(el) {
    el.parentNode.removeChild(el);
});
}