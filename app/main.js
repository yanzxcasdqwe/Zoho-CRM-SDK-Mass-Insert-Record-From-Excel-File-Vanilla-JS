ZOHO.embeddedApp.on("PageLoad",function(data){
    console.log(data);

})


function upload() {
    var files = document.getElementById('file_upload').files;
    if(files.length==0){
      alert("Please choose any file...");
      return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON(files[0]);
    }
    
    else{
        alert("Please select a valid excel file.");
    }
    document.getElementById('output').innerHTML = 'Data Pushed to CRM:'

  }
   
  function excelFileToJSON(file){
      try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function(e) {
   
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type : 'binary'
            });
            var result = {};
            workbook.SheetNames.forEach(function(sheetName) {
                var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                if (roa.length > 0) {
                    result[sheetName] = roa;
                }
            });
            let placeholder = document.querySelector("#data-output")
            let out = ''
            for (let item of result.Sheet1) {
                let ID = item.ID
                let company = item['Company Name']
                let fName = item['First Name']
                let lName = item['Last Name']
                out += `
                <tr>
                    <td class="table-dark">${ID}</td>
                    <td class="table-dark">${company}</td>
                    <td class="table-dark">${fName}</td>
                    <td class="table-dark">${lName}</td>
                </tr>
               `;
               var recordData = {
                "Company": company,
                "First_Name": fName,
                "Last_Name": lName
                }
                ZOHO.CRM.API.insertRecord({Entity:"Leads",APIData:recordData,Trigger:["workflow"]}).then(function(data){
                    console.log(data);
                });
                placeholder.innerHTML = out
            }
            document.getElementById('file_upload').value = "" 
        }
        }catch(e){
            console.error(e);
        }
  }

ZOHO.embeddedApp.init();