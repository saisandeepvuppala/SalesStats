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

function check_percentage(value1) {
	var wholesalepercentage = "";
 if (parseInt(value1) > 100 || parseInt(value1) < 0 ) {
		alert("Number should be within 0 to 100");
	}
	else {
		var x = document.getElementById("disabledInput");
		if (value1 != "") {
		x.value = (100 - parseInt(value1))+"%";
		wholesalepercentage = (100 - parseInt(value1));
		}
		else if (value1 == "" ) {
			x.value = "0%";
		}
		
		var total = document.getElementById("totalamount");
		var currentretail = document.getElementById("retailcurrentamount");
		var currentwholesale = document.getElementById("wholesalecurrentamount");
		var currenttotal = document.getElementById("totalcurrentamount");
		if (value1 != "" && total.value != "") {
			document.getElementById("rentailpending").innerText = (((parseInt(total.value))*(parseInt(value1)))/100)-(parseInt(currentretail.innerText));
			document.getElementById("totalpending").innerText = (parseInt(total.value)-parseInt(currenttotal.innerText));
			document.getElementById("wholesalepending").innerText = (((parseInt(total.value))*(parseInt(wholesalepercentage)))/100)-(parseInt(currentwholesale.innerText));
			
		}
		if (value1 == ""  || total.value == "") {
			document.getElementById("rentailpending").innerText = "";
			document.getElementById("totalpending").innerText = "";
			document.getElementById("wholesalepending").innerText = "";
			
		}
	}
}


function check_totalvalue(totalamount) {
	    var wholesalepercentage = "";
		var x = document.getElementById("disabledInput");
		wholesalepercentage = parseInt(x.value.substring(0,x.value.length-1));
		
		
		var value1 = document.getElementById("reatilpercentage").value;
		var total = totalamount;
		var currentretail = document.getElementById("retailcurrentamount");
		var currentwholesale = document.getElementById("wholesalecurrentamount");
		var currenttotal = document.getElementById("totalcurrentamount");
		if (value1 != "" && total != "") {
			document.getElementById("rentailpending").innerText = (((parseInt(total))*(parseInt(value1)))/100)-(parseInt(currentretail.innerText));
			document.getElementById("totalpending").innerText = (parseInt(total)-parseInt(currenttotal.innerText));
			document.getElementById("wholesalepending").innerText = (((parseInt(total))*(parseInt(wholesalepercentage)))/100)-(parseInt(currentwholesale.innerText));
			
		}
		if (value1 == ""  || total.value == "") {
			document.getElementById("rentailpending").innerText = "";
			document.getElementById("totalpending").innerText = "";
			document.getElementById("wholesalepending").innerText = "";
			
		}
}

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
			  document.getElementById("Ratios").innerHTML = table1 + "<div> <h3 id = 'advance'> Advance Calculations </h3>  <div class='row' id = 'row1'> <div class='col-md-3' ></div> <div class='col-md-3'  style=' padding-left: 5%; '>Retail</div> <div class='col-md-3' >Wholesale </div> <div class='col-md-3' > Total</div></div><div class='row' id = 'row2'> <div class='col-md-3' style=' padding-left: 6%; ' >Total</div> <div class='col-md-3' ><input type='Number' class='form-control' id='reatilpercentage' placeholder='Enter in %' min='0' max='100' onkeyup='check_percentage(this.value)'></div> <div class='col-md-3' > <input type='text' class='form-control'  id='disabledInput' value = '0%' disabled ></div> <div class='col-md-3' ><input type='Number' class='form-control' id='totalamount' placeholder='Enter in Num' onkeyup='check_totalvalue(this.value)'> </div></div><div class='row' id = 'row3'> <div class='col-md-3' >Current Amount </div> <div class='col-md-3' id = 'retailcurrentamount' ></div> <div class='col-md-3' id = 'wholesalecurrentamount' > </div> <div class='col-md-3' id = 'totalcurrentamount' > </div></div></div><div class='row' id = 'row4'> <div class='col-md-3' >Pending Amount </div> <div class='col-md-3' id = 'rentailpending'></div> <div class='col-md-3' id = 'wholesalepending'> </div> <div class='col-md-3' id = 'totalpending' > </div></div></div>";
			  var table2 = "<table><tr><th> Discount Type </th> <th> Percentage </th> <th> Value </th></tr><tr><td> Retail Discount </td> <td> "+((retail_discount/(retail_discount+wholesale_discount))*100).toFixed(2)+"%</td><td>"+ retail_discount.toFixed(2) +"</td></tr><tr><td>Wholesale Discount</td><td>"+((wholesale_discount/(retail_discount+wholesale_discount))*100).toFixed(2)+"%</td><td>"+wholesale_discount.toFixed(2)+" </td></tr><tr><td>Total Discount </td><td>"+(((retail_discount/(retail_discount+wholesale_discount))*100) + ((wholesale_discount/(retail_discount+wholesale_discount))*100)).toFixed(2)+"%</td><td>"+(retail_discount+wholesale_discount).toFixed(2)+"</td></tr></table>";
			  document.getElementById("Discount").innerHTML = table2;
			  var table3 = "<table><tr><th> ID </th> <th> ShopName </th> <th> Type </th> <th> SKU Count </th> </tr>";
			  for ( var i in map ) {
				  table3+= "<tr><td>"+ i + "</td><td>"+shopdetails[i][0]+"</td><td>"+shopdetails[i][1]+"</td><td>"+Object.getOwnPropertyNames(map[i]).length+"</td></tr>";
			  }
			  table3+= "</table>";
			  document.getElementById("SkuCount").innerHTML = table3;
			  document.getElementById("retailcurrentamount").innerHTML = r.toFixed(0);
			  document.getElementById("wholesalecurrentamount").innerHTML = w.toFixed(0);
			  document.getElementById("totalcurrentamount").innerHTML = (r+w).toFixed(0);
			});
        }
    }
});

 