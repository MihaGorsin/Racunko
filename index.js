function Export2Word(element, filename = ''){
    var preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
    var postHtml = "</body></html>";
    var html = preHtml+document.getElementById(element).innerHTML+postHtml;

    var blob = new Blob(['\ufeff', html], {
        type: 'application/msword'
    });
    
    // Specify link url
    var url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(html);
    
    // Specify file name
    filename = filename?filename+'.doc':'document.doc';
    
    // Create download link element
    var downloadLink = document.createElement("a");

    document.body.appendChild(downloadLink);
    
    if(navigator.msSaveOrOpenBlob ){
        navigator.msSaveOrOpenBlob(blob, filename);
    }else{
        // Create a link to the file
        downloadLink.href = url;
        
        // Setting the file name
        downloadLink.download = filename;
        
        //triggering the function
        downloadLink.click();
    }
    
    document.body.removeChild(downloadLink);
}

//Datum dneva
const d = new Date()
const curentDate = String(d.getDate()) + '. ' + String(d.getMonth() + 1) + '. ' + String(d.getFullYear())
document.getElementById('datumDanes').innerHTML= curentDate

//Event listeners
document.getElementById('gumbVnos').addEventListener('click', izpolni);
document.getElementById('gumbPrenos').addEventListener('click', download);

function pridobiPodatke(){
    var narocnik = document.getElementById('narocnikVnos').value;
    var storitev = document.getElementById('storitevVnos').value;
    var cena = document.getElementById('cenaVnos').value;
    return {narocnik, storitev, cena}; // Return an object containing the values
}

function izpolni(){
    var {narocnik, storitev, cena} = pridobiPodatke();
    document.getElementById('narocnikIzpis').innerHTML = narocnik
    document.getElementById('storitevIzpis').innerHTML = storitev
    document.getElementById('cenaIzpis').innerHTML = cena

    document.getElementById('storitevIzpis2').innerHTML = storitev
    document.getElementById('cenaIzpis2').innerHTML = cena
}

function download(){
    Export2Word('exportContent', 'Racun - ' + document.getElementById('narocnikVnos').value + '.docx')
}