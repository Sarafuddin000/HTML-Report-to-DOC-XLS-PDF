--I have a HTML Report in Oracle apex. Now i need to export the that HTML report into Microsoft Word Document. To achieve this functionality you need to add some JavaScript functions in Apex page.

--Firstly, Write HTML Report using Table, try to use les div as possible for this purpose. Use Div ID For region to be PRINT and table ID to be export as document.

    begin
        htp.p('
            <div id="source">
            <table id="main-table" class="main-page">
    
    
        ');
        FOR M IN m1
        loop
    htp.p( '<tr class="page" style="width:100%"> <td style="width:100%;">');
                for xxx in (........................


--2. Use "jobCard" as Region's static ID

--3. Add page level inline CSS as required
          
        .page{
             size:  8.3in 11.7in ;
          
          }
          
          .main-page{
              width: 100%;
          }
          
          #data-table{
              width: 100%;
          }
          
          .t-center{
              text-align: center;
              font-size: 9pt;
          
          }
          
          .td{
              text-align: center;
              font-size: 9pt;
          }

--4. For PRINT feature write function in Page level function and Global variable section
          function printdiv(id)
          			{
          			var headstr = "<html><head><title></title></head><body>";
          			var footstr = "</body>";
          			var newstr = document.all.item(id).innerHTML;
          			var oldstr = document.body.innerHTML;
          			document.body.innerHTML = headstr+newstr+footstr;
          			window.print();
          			document.body.innerHTML = oldstr;
          			return false;
          			}

--5. For export as DOC  feature write function in Page level function and Global variable section
            function Export2Word( filename ){
                var preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
                var postHtml = "</body></html>";
                var html = preHtml+document.getElementById("source").innerHTML+postHtml;
            
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

--6. For export as XLSX  feature write function in Page level function and Global variable section

            
          function Export2Excel(filename) {
              var preHtml = "<html xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Excel</title></head><body>";
              var postHtml = "</body></html>";
              var html = preHtml + document.getElementById("source").innerHTML + postHtml;
          
              var blob = new Blob(['\ufeff', html], {
                  type: 'application/vnd.ms-excel'
              });
          
              // Specify link url
              var url = 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(html);
          
              // Specify file name
              filename = filename ? filename + '.xls' : 'document.xls';
          
              // Create download link element
              var downloadLink = document.createElement("a");
          
              document.body.appendChild(downloadLink);
          
              if (navigator.msSaveOrOpenBlob) {
                  navigator.msSaveOrOpenBlob(blob, filename);
              } else {
                  // Create a link to the file
                  downloadLink.href = url;
          
                  // Setting the file name
                  downloadLink.download = filename;
          
                  // Triggering the function
                  downloadLink.click();
              }
          
              document.body.removeChild(downloadLink);
          }


--7.  Now create button for PRINT, DOCX, XLSX and use Execute Javascript code for respective purpose as below:
        PRINT    :   printdiv('source')
        DOCS     :   Export2Word('jobCard')
        XLSX     :   Export2Excel('jobCard')

--8.  Great, Now you can export or print your html report directly and modify after download.
