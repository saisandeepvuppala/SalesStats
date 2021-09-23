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
			document.getElementById("rentailpending").innerText = ((((parseInt(total.value))*(parseInt(value1)))/100)-(parseInt(currentretail.innerText))).toFixed(0);
			document.getElementById("totalpending").innerText = ((parseInt(total.value)-parseInt(currenttotal.innerText))).toFixed(0);
			document.getElementById("wholesalepending").innerText = ((((parseInt(total.value))*(parseInt(wholesalepercentage)))/100)-(parseInt(currentwholesale.innerText))).toFixed(0);
			
		}
		if (value1 == ""  || total == "") {
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
			document.getElementById("rentailpending").innerText = ((((parseInt(total))*(parseInt(value1)))/100)-(parseInt(currentretail.innerText))).toFixed(0);
			document.getElementById("totalpending").innerText = ((parseInt(total)-parseInt(currenttotal.innerText))).toFixed(0);
			document.getElementById("wholesalepending").innerText = ((((parseInt(total))*(parseInt(wholesalepercentage)))/100)-(parseInt(currentwholesale.innerText))).toFixed(0);
			
		}
		if (value1 == ""  || total == "") {
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
			var retail_maxx = 0, retail_doritos = 0, retail_wafer = 0;
			var retail_maxx_discount = 0; retail_doritos_discount = 0; retail_wafer_discount = 0;
			var wholesale_maxx = 0, wholesale_doritos = 0, wholesale_wafer = 0;
			var wholesale_maxx_discount = 0; wholesale_doritos_discount = 0; wholesale_wafer_discount = 0;
			var total_maxx = 0; total_doritos = 0; total_wafer = 0;
			var total_maxx_discount = 0; total_doritos_discount = 0; total_wafer_discount = 0;
			var map = {};
			var shopdetails = {};
			var individual_Shop_value = {};
			  for (var k in rowObject) { 
				if (parseInt(k) >= 6) { 
					
					if (parseInt(k) == 130 ) {
						var a = 0 ;
					}
					if(rowObject[k].__EMPTY_9 != null && rowObject[k].__EMPTY_9.endsWith("Retail")){
						
						if (rowObject[k].__EMPTY_19 != null && rowObject[k].__EMPTY_19.endsWith("MAXX")) {
							retail_maxx = retail_maxx + rowObject[k].__EMPTY_37;
							retail_maxx_discount = retail_maxx_discount + rowObject[k].__EMPTY_34;
						} else if (rowObject[k].__EMPTY_19 != null && rowObject[k].__EMPTY_19.endsWith("Doritos")) {
							retail_doritos = retail_doritos + rowObject[k].__EMPTY_37;
							retail_doritos_discount = retail_doritos_discount + rowObject[k].__EMPTY_34;
						} else if (rowObject[k].__EMPTY_19 != null && rowObject[k].__EMPTY_19.endsWith("Wafer Style")) {
							retail_wafer = retail_wafer + rowObject[k].__EMPTY_37;
							retail_wafer_discount = retail_wafer_discount + rowObject[k].__EMPTY_34;
						}
						r+= rowObject[k].__EMPTY_37;
						if (individual_Shop_value[rowObject[k].__EMPTY_5] == null){
							individual_Shop_value[rowObject[k].__EMPTY_5] = (rowObject[k].__EMPTY_37);
						} else {
							individual_Shop_value[rowObject[k].__EMPTY_5] += (rowObject[k].__EMPTY_37);
						}
						 
						retail_discount+= rowObject[k].__EMPTY_34;
						if (map[rowObject[k].__EMPTY_5] != null) {
							var value = {};
								if (map[rowObject[k].__EMPTY_5][rowObject[k].__EMPTY_31] == null) {
									value = map[rowObject[k].__EMPTY_5];
									value[rowObject[k].__EMPTY_31] = 1;
									map[rowObject[k].__EMPTY_5]  = value;
							}
						}
						else {
							var value = {};
							value[rowObject[k].__EMPTY_31] = 1;
							map[rowObject[k].__EMPTY_5]  = value;
						}
						
						if (shopdetails[rowObject[k].__EMPTY_5] == null) {
							var details = [];
							details.push(rowObject[k].__EMPTY_6);
							details.push(rowObject[k].__EMPTY_9);
							shopdetails[rowObject[k].__EMPTY_5] = details;
						}
					}  else if (rowObject[k].__EMPTY_9 != null && rowObject[k].__EMPTY_9.endsWith("WHOLESALE"))  {
						
						if (rowObject[k].__EMPTY_19 != null && rowObject[k].__EMPTY_19.endsWith("MAXX")) {
							wholesale_maxx = wholesale_maxx + rowObject[k].__EMPTY_37;
							wholesale_maxx_discount = wholesale_maxx_discount + rowObject[k].__EMPTY_34;
						} else if (rowObject[k].__EMPTY_19 != null && rowObject[k].__EMPTY_19.endsWith("Doritos")) {
							wholesale_doritos = wholesale_doritos + rowObject[k].__EMPTY_37;
							wholesale_doritos_discount = wholesale_doritos_discount + rowObject[k].__EMPTY_34;
						} else if (rowObject[k].__EMPTY_19 != null && rowObject[k].__EMPTY_19.endsWith("Wafer Style")) {
							wholesale_wafer = wholesale_wafer + rowObject[k].__EMPTY_37;
							wholesale_wafer_discount = wholesale_wafer_discount + rowObject[k].__EMPTY_34;
						}
						if (individual_Shop_value[rowObject[k].__EMPTY_5] == null){
							individual_Shop_value[rowObject[k].__EMPTY_5] = (rowObject[k].__EMPTY_37);
						} else {
							individual_Shop_value[rowObject[k].__EMPTY_5] += (rowObject[k].__EMPTY_37);
						}
						w+= rowObject[k].__EMPTY_37;
						wholesale_discount+= rowObject[k].__EMPTY_34;
						if (map[rowObject[k].__EMPTY_5] != null) {
							var value = {};
								if (map[rowObject[k].__EMPTY_5][rowObject[k].__EMPTY_31] == null) {
									value = map[rowObject[k].__EMPTY_5];
									value[rowObject[k].__EMPTY_31] = 1;
									map[rowObject[k].__EMPTY_5]  = value;
							}
						}
						else {
							var value = {};
							value[rowObject[k].__EMPTY_31] = 1;
							map[rowObject[k].__EMPTY_5]  = value;
						}
						
						if (shopdetails[rowObject[k].__EMPTY_5] == null) {
							var details = [];
							details.push(rowObject[k].__EMPTY_6);
							details.push(rowObject[k].__EMPTY_9);
							shopdetails[rowObject[k].__EMPTY_5] = details;
						}
					}
				}
				total_maxx = retail_maxx + wholesale_maxx;
				total_wafer = retail_wafer + wholesale_wafer;
				total_doritos = retail_doritos + wholesale_doritos;
				total_maxx_discount  = retail_maxx_discount + wholesale_maxx_discount;
				total_wafer_discount = retail_wafer_discount + wholesale_wafer_discount;
				total_doritos_discount = retail_doritos_discount + wholesale_doritos_discount;
				
			  }
			  document.getElementById("Title").innerHTML = "Sales Stats";
			  document.getElementById("From").innerHTML = "Date From : " + rowObject[1].__EMPTY_1 ;
			  document.getElementById("To").innerHTML = " Date To : " + rowObject[2].__EMPTY_1;
			  
			  var table2 = "<table style='margin-left: 40px; '><tr><th> Discount Type </th> <th> Percentage </th> <th> Value </th></tr><tr><td> Retail Discount </td> <td> "+((retail_discount/(retail_discount+wholesale_discount))*100).toFixed(2)+"%</td><td>"+ retail_discount.toFixed(0) +"</td></tr><tr><td>Wholesale Discount</td><td>"+((wholesale_discount/(retail_discount+wholesale_discount))*100).toFixed(2)+"%</td><td>"+wholesale_discount.toFixed(0)+" </td></tr><tr><td>Total Discount </td><td>"+(((retail_discount/(retail_discount+wholesale_discount))*100) + ((wholesale_discount/(retail_discount+wholesale_discount))*100)).toFixed(2)+"%</td><td>"+(retail_discount+wholesale_discount).toFixed(0)+"</td></tr></table>          <table style='margin-left: 40px; '><tr><th> Name </th> <th> Discount  </th> <th> Value </th></tr><tr><td> Total Maxx </td> <td> "+ (total_maxx_discount).toFixed(2)+"</td><td>"+ total_maxx.toFixed(0) +"</td></tr><tr><td>Total Wafer Style</td><td>"+(total_wafer_discount).toFixed(2)+"</td><td>"+total_wafer.toFixed(0)+" </td></tr><tr><td>Total Doritos </td><td>"+(total_doritos_discount).toFixed(2)+"</td><td>"+(total_doritos).toFixed(0)+"</td></tr></table>";
			  document.getElementById("Discount").innerHTML = table2;
			  var table3 = "<table class='table-sortable' style='margin-left: -2%;'><thead><tr><th> ID </th> <th> ShopName </th> <th> Type </th> <th> SKU_Count </th> <th> Value </th> </tr></thead><tbody>";
			  var retailmap = {};
			  var wholesalemap = {};
			  for (var i in map) {
			  	if (shopdetails[i][1] == "Retail") {
			  		retailmap[i] = Object.getOwnPropertyNames(map[i]).length;
			  	} else {
			  		wholesalemap[i] = Object.getOwnPropertyNames(map[i]).length;
			  	}
			  }
			  var sortretailmap = Object.entries(retailmap).sort((a,b)=>a[1]-b[1]);
			  var sortwholesalemap = Object.entries(wholesalemap).sort((a,b)=>a[1]-b[1]);
			  var retail_less_sku_count = 0;
			  var wholesale_less_sku_count = 0;
			  var total_retail_shops = 0;
			  var total_wholesale_shops = 0;

			  for (var i in sortretailmap) {
			  	var skunotmet = "";
			  		if (sortretailmap[i][1]< 8) {
			  			skunotmet = "style= 'background-color: red;' ";
			  			retail_less_sku_count++;
			  		} 
			  		total_retail_shops++;
				  table3+= "<tr style= 'background-color: yellow;'><td>"+ sortretailmap[i][0] + "</td><td>"+shopdetails[sortretailmap[i][0]][0]+"</td><td>"+shopdetails[sortretailmap[i][0]][1]+"</td><td "+skunotmet+">"+ sortretailmap[i][1] +"</td> <td> " + individual_Shop_value[sortretailmap[i][0]].toFixed(0)+ "</td></tr>";
			  }

			  for (var i in sortwholesalemap) {
			  	var skunotmet = "";
			  		if (sortwholesalemap[i][1]< 6) {
			  			skunotmet = "style= 'background-color: red;' ";
			  			wholesale_less_sku_count++;
			  		} 
			  		total_wholesale_shops++;
				  table3+= "<tr style= 'background-color: rgb(0 123 255 / 68%); '><td>"+ sortwholesalemap[i][0] + "</td><td>"+shopdetails[sortwholesalemap[i][0]][0]+"</td><td>"+shopdetails[sortwholesalemap[i][0]][1]+"</td><td "+skunotmet+">"+ sortwholesalemap[i][1] +"</td> <td> " + individual_Shop_value[sortwholesalemap[i][0]].toFixed(0) +"</td></tr>";
			  }

			  
			  table3+= "</tbody></table>";
			  var All_Retail_Shops = 119;
			  var All_Wholesale_Shops = 28;
			  
			  var table1 = "<table style='margin-left: 2px;margin-right: 0px;width: 115%; '><tr><th> Type </th> <th> Percentage </th> <th> Value </th> <th> Shops Billed  </th> <th> Shops Unbilled </th> <th> Less_SKU_Count</th></tr><tr><td> Retail </td> <td> "+((r/(r+w))*100).toFixed(2)+"%</td><td>"+ r.toFixed(0) +"</td><td> " + total_retail_shops + " </td> <td> " + (All_Retail_Shops - total_retail_shops) + " ("+(((All_Retail_Shops - total_retail_shops)/(All_Retail_Shops))*100).toFixed(0)+"%) </td> <td> " + retail_less_sku_count  + " (Below 8) </td></tr><tr><td>Wholesale</td><td>"+((w/(r+w))*100).toFixed(2)+"%</td><td>"+w.toFixed(0)+" </td><td>"+total_wholesale_shops+"</td> <td> " + (All_Wholesale_Shops - total_wholesale_shops)+ " ("+(((All_Wholesale_Shops - total_wholesale_shops)/(All_Wholesale_Shops))*100).toFixed(0)+"%)</td><td>"+ wholesale_less_sku_count +" (Below 6) </td></tr><tr><td>Total Trade Price </td><td>"+(((r/(r+w))*100) + ((w/(r+w))*100)).toFixed(2)+"%</td><td>"+(r+w).toFixed(0)+"</td><td>"+ (total_retail_shops+total_wholesale_shops).toFixed(0)+"</td> <td> " + ((All_Retail_Shops + All_Wholesale_Shops) - (total_retail_shops+total_wholesale_shops)).toFixed(0)+ " ("+(((((All_Retail_Shops + All_Wholesale_Shops) - (total_retail_shops+total_wholesale_shops)))/(All_Retail_Shops+All_Wholesale_Shops))*100).toFixed(0)+"%)</td><td>"+ (retail_less_sku_count+wholesale_less_sku_count).toFixed(0) +"</td></tr></table>";
			  document.getElementById("Ratios").innerHTML = table1 + "<div> <h3 id = 'advance' style= ' margin-top: 34%; '> Advance Calculations </h3>  <div class='row' id = 'row1'> <div class='col-md-3' ></div> <div class='col-md-3'  style=' padding-left: 5%; '>Retail</div> <div class='col-md-3' >Wholesale </div> <div class='col-md-3' > Total</div></div><div class='row' id = 'row2'> <div class='col-md-3' style=' padding-left: 6%; ' >Total</div> <div class='col-md-3' ><input type='Number' class='form-control' id='reatilpercentage' placeholder='Enter in %' min='0' max='100' onkeyup='check_percentage(this.value)'></div> <div class='col-md-3' > <input type='text' class='form-control'  id='disabledInput' value = '0%' disabled ></div> <div class='col-md-3' ><input type='Number' class='form-control' id='totalamount' placeholder='Enter in Num' onkeyup='check_totalvalue(this.value)'> </div></div><div class='row' id = 'row3'> <div class='col-md-3' >Current Amount </div> <div class='col-md-3' id = 'retailcurrentamount' ></div> <div class='col-md-3' id = 'wholesalecurrentamount' > </div> <div class='col-md-3' id = 'totalcurrentamount' > </div></div></div><div class='row' id = 'row4'> <div class='col-md-3' >Pending Amount </div> <div class='col-md-3' id = 'rentailpending' style=' padding-top: 1.5%; '></div> <div class='col-md-3' id = 'wholesalepending' style=' padding-top: 1.5%; '> </div> <div class='col-md-3' id = 'totalpending' style=' padding-top: 1.5%;'> </div></div></div>";

			  document.getElementById("SkuCount").innerHTML = table3;
			  document.getElementById("retailcurrentamount").innerHTML = r.toFixed(0) + "<span style='font-size: 12px; background-color: rgb(0 123 255 / 50%); '> ("+((r/(r+w))*100).toFixed(2)+"%)&nbsp</span>";
			  document.getElementById("wholesalecurrentamount").innerHTML = w.toFixed(0) + "<span style='font-size: 12px; background-color: rgb(0 123 255 / 50%); '> ("+((w/(r+w))*100).toFixed(2)+"%)&nbsp</span>";
			  document.getElementById("totalcurrentamount").innerHTML = (r+w).toFixed(0) + "<span style='font-size: 12px; background-color: rgb(0 123 255 / 50%); '> ("+(((r/(r+w))*100) + ((w/(r+w))*100)).toFixed(2)+"%)&nbsp</span>";


			});
        }
    }
});


/**
 * Sorts a HTML table.
 * 
 * @param {HTMLTableElement} table The table to sort
 * @param {number} column The index of the column to sort
 * @param {boolean} asc Determines if the sorting will be in ascending
 */
function sortTableByColumn(table, column, asc = true) {
    const dirModifier = asc ? 1 : -1;
    const tBody = table.tBodies[0];
    const rows = Array.from(tBody.querySelectorAll("tr"));

    // Sort each row
    const sortedRows = rows.sort((a, b) => {
        const aColText = a.querySelector(`td:nth-child(${ column + 1 })`).textContent.trim();
        const bColText = b.querySelector(`td:nth-child(${ column + 1 })`).textContent.trim();

        return aColText > bColText ? (1 * dirModifier) : (-1 * dirModifier);
    });

    // Remove all existing TRs from the table
    while (tBody.firstChild) {
        tBody.removeChild(tBody.firstChild);
    }

    // Re-add the newly sorted rows
    tBody.append(...sortedRows);

    // Remember how the column is currently sorted
    table.querySelectorAll("th").forEach(th => th.classList.remove("th-sort-asc", "th-sort-desc"));
    table.querySelector(`th:nth-child(${ column + 1})`).classList.toggle("th-sort-asc", asc);
    table.querySelector(`th:nth-child(${ column + 1})`).classList.toggle("th-sort-desc", !asc);
}

document.querySelectorAll(".table-sortable th").forEach(headerCell => {
    headerCell.addEventListener("click", () => {
        const tableElement = headerCell.parentElement.parentElement.parentElement;
        const headerIndex = Array.prototype.indexOf.call(headerCell.parentElement.children, headerCell);
        const currentIsAscending = headerCell.classList.contains("th-sort-asc");

        sortTableByColumn(tableElement, headerIndex, !currentIsAscending);
    });
}); 

 
