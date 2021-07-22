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
			var map = {};
			var shopdetails = {};
			  for (var k in rowObject) { 
				if (parseInt(k) >= 6) { 
					if(rowObject[k].__EMPTY_9 != null && rowObject[k].__EMPTY_9.endsWith("Retail")){
						r+= rowObject[k].__EMPTY_37;
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
			  }
			  document.getElementById("Title").innerHTML = "Sales Stats";
			  document.getElementById("From").innerHTML = "Date From : " + rowObject[1].__EMPTY_1 ;
			  document.getElementById("To").innerHTML = " Date To : " + rowObject[2].__EMPTY_1;
			  
			  var table2 = "<table style='margin-left: 66px; '><tr><th> Discount Type </th> <th> Percentage </th> <th> Value </th></tr><tr><td> Retail Discount </td> <td> "+((retail_discount/(retail_discount+wholesale_discount))*100).toFixed(2)+"%</td><td>"+ retail_discount.toFixed(0) +"</td></tr><tr><td>Wholesale Discount</td><td>"+((wholesale_discount/(retail_discount+wholesale_discount))*100).toFixed(2)+"%</td><td>"+wholesale_discount.toFixed(0)+" </td></tr><tr><td>Total Discount </td><td>"+(((retail_discount/(retail_discount+wholesale_discount))*100) + ((wholesale_discount/(retail_discount+wholesale_discount))*100)).toFixed(2)+"%</td><td>"+(retail_discount+wholesale_discount).toFixed(0)+"</td></tr></table>";
			  document.getElementById("Discount").innerHTML = table2;
			  var table3 = "<table class='table-sortable' style='margin-left: 2%;'><thead><tr><th> ID </th> <th> ShopName </th> <th> Type </th> <th> SKU_Count </th> </tr></thead><tbody>";
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
				  table3+= "<tr style= 'background-color: yellow;'><td>"+ sortretailmap[i][0] + "</td><td>"+shopdetails[sortretailmap[i][0]][0]+"</td><td>"+shopdetails[sortretailmap[i][0]][1]+"</td><td "+skunotmet+">"+ sortretailmap[i][1] +"</td></tr>";
			  }

			  for (var i in sortwholesalemap) {
			  	var skunotmet = "";
			  		if (sortwholesalemap[i][1]< 6) {
			  			skunotmet = "style= 'background-color: red;' ";
			  			wholesale_less_sku_count++;
			  		} 
			  		total_wholesale_shops++;
				  table3+= "<tr style= 'background-color: rgb(0 123 255 / 68%); '><td>"+ sortwholesalemap[i][0] + "</td><td>"+shopdetails[sortwholesalemap[i][0]][0]+"</td><td>"+shopdetails[sortwholesalemap[i][0]][1]+"</td><td "+skunotmet+">"+ sortwholesalemap[i][1] +"</td></tr>";
			  }

			  
			  table3+= "</tbody></table>";

			  var table1 = "<table style='margin-left: 13px;margin-right: 0px;width: 115%; '><tr><th> Type </th> <th> Percentage </th> <th> Value </th> <th> Shops Billed Count </th> <th> Less_SKU_Count</th></tr><tr><td> Retail </td> <td> "+((r/(r+w))*100).toFixed(2)+"%</td><td>"+ r.toFixed(0) +"</td><td> " + total_retail_shops + " </td><td> " + retail_less_sku_count  + " (Below 8) </td></tr><tr><td>Wholesale</td><td>"+((w/(r+w))*100).toFixed(2)+"%</td><td>"+w.toFixed(0)+" </td><td>"+total_wholesale_shops+"</td><td>"+ wholesale_less_sku_count +" (Below 6) </td></tr><tr><td>Total Trade Price </td><td>"+(((r/(r+w))*100) + ((w/(r+w))*100)).toFixed(2)+"%</td><td>"+(r+w).toFixed(0)+"</td><td>"+ (total_retail_shops+total_wholesale_shops).toFixed(0)+"</td><td>"+ (retail_less_sku_count+wholesale_less_sku_count).toFixed(0) +"</td></tr></table>";
			  document.getElementById("Ratios").innerHTML = table1 + "<div> <h3 id = 'advance'> Advance Calculations </h3>  <div class='row' id = 'row1'> <div class='col-md-3' ></div> <div class='col-md-3'  style=' padding-left: 5%; '>Retail</div> <div class='col-md-3' >Wholesale </div> <div class='col-md-3' > Total</div></div><div class='row' id = 'row2'> <div class='col-md-3' style=' padding-left: 6%; ' >Total</div> <div class='col-md-3' ><input type='Number' class='form-control' id='reatilpercentage' placeholder='Enter in %' min='0' max='100' onkeyup='check_percentage(this.value)'></div> <div class='col-md-3' > <input type='text' class='form-control'  id='disabledInput' value = '0%' disabled ></div> <div class='col-md-3' ><input type='Number' class='form-control' id='totalamount' placeholder='Enter in Num' onkeyup='check_totalvalue(this.value)'> </div></div><div class='row' id = 'row3'> <div class='col-md-3' >Current Amount </div> <div class='col-md-3' id = 'retailcurrentamount' ></div> <div class='col-md-3' id = 'wholesalecurrentamount' > </div> <div class='col-md-3' id = 'totalcurrentamount' > </div></div></div><div class='row' id = 'row4'> <div class='col-md-3' >Pending Amount </div> <div class='col-md-3' id = 'rentailpending' style=' padding-top: 1.5%; '></div> <div class='col-md-3' id = 'wholesalepending' style=' padding-top: 1.5%; '> </div> <div class='col-md-3' id = 'totalpending' style=' padding-top: 1.5%;'> </div></div></div>";

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

 