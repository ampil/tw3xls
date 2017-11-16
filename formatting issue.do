** putexcel/mata question

//set trace on
//cd "c:/temp"
local sheet = "Data"
local using = "Example1"
local ext = "xls"					//setting it to "xlsx" fixes formatting but increases the execution time
cap erase "`using'.`ext'"

*Create matrix to export
mata: A = ((1,2,3,4,5,6,7) \ (2,2,2,2,2,2,2) \ (8,8,8,8,8,8,8) \ (3,3,3,3,3,3,3) \ (3,3,3,3,3,3,3)  )
mata: st_local("cols", strofreal(cols(A)))
mata: st_local("rows", strofreal(rows(A)))
mata: st_local("rcolf", numtobase26(`cols'))		//Last column name in the Excel notation

* Create an Excel workbook
mata:	b = xl()
mata: b.create_book("`using'", "`sheet'", "`ext'")

* export i tables to Excel; when increasing the number of tables (i), more cells left unformatted
forval i = 1/6 {

	* set positions to paste the data into a spreadsheet
	local gap = 5 
	local startrow  = (`i' - 1) * (`rows' + `gap') + 1
	local secondrow = `startrow' + 1	
	local thirdrow  = `startrow' + 2
	local fourthrow = `startrow' + 3
	local fifthrow  = `startrow' + 4
	local lastrow   = `startrow' + `rows' + `gap' - 2	

	* save row positions into mata vectors
	mata: row1 = strtoreal(st_local("startrow"))
	mata: row2 = strtoreal(st_local("secondrow"))
	mata: row3 = strtoreal(st_local("thirdrow"))
	mata: row4 = strtoreal(st_local("fourthrow"))
	mata: row5 = strtoreal(st_local("fifthrow"))
	mata: rowN = strtoreal(st_local("lastrow"))
	
	* paste the matrix and some strings to test formatting: borders and alignment
	mata:	b.put_number(row5, 1, A)
	mata:	b.put_string(row2, 1, "Center text1")
	mata:	b.put_string(row3, `cols', "Center text2")
	mata:	b.put_string(row2, 2, "Center text3")
	
	* merge the first range (appopriate 
	//qui putexcel A`secondrow':A`fourthrow' , merge hcenter txtwrap
	mata: rows_vector = (row2, row4)
	mata: cols_vector = (1, 1)
	mata: b.set_sheet_merge("`sheet'", rows_vector, cols_vector)
	mata: b.set_vertical_align(rows_vector, cols_vector, "center")
	mata: b.set_horizontal_align(rows_vector, cols_vector, "center")
	mata: b.set_text_wrap(rows_vector, cols_vector,"on")
	
	* merge the second range
	//qui putexcel B`secondrow':`rcolf'`secondrow', merge hcenter txtwrap
	mata: rows_vector = (row2, row2)
	mata: cols_vector = (2, `cols')
	mata: b.set_sheet_merge("`sheet'", rows_vector, cols_vector)
	mata: b.set_horizontal_align(rows_vector, cols_vector, "center")
	
	* merge the third range
	//qui putexcel `rcolf'`thirdrow':`rcolf'`fourthrow', merge hcenter txtwrap
	mata: rows_vector = (row3, row4)
	mata: cols_vector = (`cols', `cols')
	mata: b.set_sheet_merge("`sheet'", rows_vector, cols_vector)
	mata: b.set_horizontal_align(rows_vector, cols_vector, "center")
	mata: b.set_text_wrap(rows_vector, cols_vector,"on")
	
	di "Iteraction `i':"
	di "Rows to set borders: `secondrow':`lastrow'"
	di "Columns to set borders: 1:`cols'"
	di ""
	
	* draw borders and set vertical alignment everywhere
	//qui putexcel A`secondrow':`rcolf'`lastrow'  , border(all) vcenter 
	mata: rows_vector = (row2, rowN)
	mata: cols_vector = (1, `cols')
	mata: b.set_vertical_align(rows_vector, cols_vector, "center")
	mata: b.set_border(rows_vector, cols_vector, "thin")
	
}

mata: b.close_book()
di in yellow "Output written to {browse `using'.`ext'}"
