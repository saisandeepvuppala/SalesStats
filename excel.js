let selectedFile;
console.log(window.XLSX);
document.getElementById('input').addEventListener("change", (event) => {
    selectedFile = event.target.files[0];
})

let data=[{
    "name":"Sandeep",
    "data":"scd",
    "abc":"sdef"
}]

document.getElementById('button').addEventListener("click", () => {
    XLSX.utils.json_to_sheet(data, 'out.xlsx');
    if(selectedFile){
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);
        fileReader.onload = (event)=>{
         let data = event.target.result;
         let workbook = XLSX.read(data,{type:"binary"});
         console.log(workbook);
         workbook.SheetNames.forEach(sheet => {
              let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
              console.log(rowObject);
			var r = 0, w = 0, retail_discount = 0, wholesale_discount = 0;
			var map = {};
			var shopdetails = {};
			  for (var k in rowObject) { 
				if (parseInt(k) >= 6) { 
					if(rowObject[k]._11.endsWith("Retail")){
						r+= rowObject[k].__EMPTY_6;
						retail_discount+= rowObject[k].__EMPTY_3;
						if (map[rowObject[k]._8] != null) {
							var value = {};
								if (map[rowObject[k]._8][rowObject[k]._33] == null) {
									value = map[rowObject[k]._8];
									value[rowObject[k]._33] = 1;
									map[rowObject[k]._8]  = value;
							}
						}
						else {
							var value = {};
							value[rowObject[k]._33] = 1;
							map[rowObject[k]._8]  = value;
						}
						
						if (shopdetails[rowObject[k]._8] == null) {
							var details = [];
							details.push(rowObject[k]._9);
							details.push(rowObject[k]._11);
							shopdetails[rowObject[k]._8] = details;
						}
					}  else if (rowObject[k]._11.endsWith("WHOLESALE"))  {
						w+= rowObject[k].__EMPTY_6;
						wholesale_discount+= rowObject[k].__EMPTY_3;
						if (map[rowObject[k]._8] != null) {
							var value = {};
								if (map[rowObject[k]._8][rowObject[k]._33] == null) {
									value = map[rowObject[k]._8];
									value[rowObject[k]._33] = 1;
									map[rowObject[k]._8]  = value;
							}
						}
						else {
							var value = {};
							value[rowObject[k]._33] = 1;
							map[rowObject[k]._8]  = value;
						}
						
						if (shopdetails[rowObject[k]._8] == null) {
							var details = [];
							details.push(rowObject[k]._9);
							details.push(rowObject[k]._11);
							shopdetails[rowObject[k]._8] = details;
						}
					}
				}
			  }
			  document.getElementById("Title").innerHTML = "Sales Stats";
			  document.getElementById("From").innerHTML = "Date From : " + rowObject[1]._1 ;
			  document.getElementById("To").innerHTML = " Date To : " + rowObject[2]._1;
			  var table1 = "<table><tr><th> Type </th> <th> Percentage </th> <th> Value </th></tr><tr><td> Retail </td> <td> "+((r/(r+w))*100).toFixed(2)+"%</td><td>"+ r.toFixed(2) +"</td></tr><tr><td>Wholesale</td><td>"+((w/(r+w))*100).toFixed(2)+"%</td><td>"+w.toFixed(2)+" </td></tr><tr><td>Total Trade Price </td><td>"+(((r/(r+w))*100) + ((w/(r+w))*100)).toFixed(2)+"%</td><td>"+(r+w).toFixed(2)+"</td></tr></table>";
			  document.getElementById("Ratios").innerHTML = table1;
			  var table2 = "<table><tr><th> Discount Type </th> <th> Percentage </th> <th> Value </th></tr><tr><td> Retail Discount </td> <td> "+((retail_discount/(retail_discount+wholesale_discount))*100).toFixed(2)+"%</td><td>"+ retail_discount.toFixed(2) +"</td></tr><tr><td>Wholesale Discount</td><td>"+((wholesale_discount/(retail_discount+wholesale_discount))*100).toFixed(2)+"%</td><td>"+wholesale_discount.toFixed(2)+" </td></tr><tr><td>Total Discount </td><td>"+(((retail_discount/(retail_discount+wholesale_discount))*100) + ((wholesale_discount/(retail_discount+wholesale_discount))*100)).toFixed(2)+"%</td><td>"+(retail_discount+wholesale_discount).toFixed(2)+"</td></tr></table>";
			  document.getElementById("Discount").innerHTML = table2;
			  var table3 = "<table><tr><th> ID </th> <th> ShopName </th> <th> Type </th> <th> SKU Count </th> </tr>";
			  for ( var i in map ) {
				  table3+= "<tr><td>"+ i + "</td><td>"+shopdetails[i][0]+"</td><td>"+shopdetails[i][1]+"</td><td>"+Object.getOwnPropertyNames(map[i]).length+"</td></tr>";
			  }
			  table3+= "</table>";
			  document.getElementById("SkuCount").innerHTML = table3;
			});
        }
    }
});

 