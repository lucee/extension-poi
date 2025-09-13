component {
	/*
INSTALLATION:
1. extract the latest zip (or the one you like) from here https://archive.apache.org/dist/poi/release/bin/ locally
2.copy the jars in the zip to "lib/ext" in your Lucee installation  (same place where the lucee.jar is) or define it with your javaSettings in the application.cfc
3.POI.cfc copy this POI to the foldrr where you keep your  components
*/



	// jar version 3.15.0
	// load classes
	variables.HorizontalAlignment=POI::loadPoi("org.apache.poi.ss.usermodel.HorizontalAlignment");
	variables.BorderStyle=POI::loadPoi("org.apache.poi.ss.usermodel.BorderStyle");
	variables.CellStyle=POI::loadPoi("org.apache.poi.ss.usermodel.CellStyle");
	variables.Font=POI::loadPoi("org.apache.poi.ss.usermodel.Font");
	variables.XSSFColor=POI::loadPoi("org.apache.poi.xssf.usermodel.XSSFColor");
	variables.ColorCaster=POI::loadSystem("lucee.commons.color.ColorCaster");
	variables.CellUtil=POI::loadSystem("org.apache.poi.ss.util.CellUtil");
	variables.StringUtils=createObject('java','org.apache.commons.lang.StringUtils');
	variables.FillPatternType=createObject('java','org.apache.poi.ss.usermodel.FillPatternType');
	






	variables.dateTimeFormat="dd.mm.yyyy h:mm";//"m/d/yy h:mm";
	variables.timeFormat="hh:mm";//"m/d/yy h:mm";
	variables.dateFormat="dd.mm.yyyy";//"m/d/yy h:mm";

	// styles
	variables.dateStyles={};
	variables.cellStyles={};
	variables.rowColStyles={};

	variables.fonts={};
	variables.colors={};


	private void function init(required poi,file='', required boolean isNew) localmode=true {
		variables.poi=arguments.poi;
		variables.file=arguments.file;
		variables.isNew=arguments.isNew;
	}

	/**
	* creates a POI for an existing spreadsheet
	* @path path to the existing spreadsheet
	* @password password for the spreadsheet
	*/
	public static function getInstance(required string path, string password) {
		var FileInputStream=loadSystem("java.io.FileInputStream");
		var WorkbookFactory=loadPoi("org.apache.poi.ss.usermodel.WorkbookFactory");

		// check path
		path=expandPath(path);
		
		try {
			var file = expandPath(path);
			var fis = FileInputStream.init(file);
			if(isNull(password) || len(password)==0)
				local.wb=WorkbookFactory.create(fis);
			else
				local.wb=WorkbookFactory.create(fis,password);
			return new POI(wb,file,false); 
		}
		finally {
			if(!isNull(fis))fis.close();
		}
	}


	/**
	* creates a POI for an existing spreadsheet
	* @pathOrType path to the spreadsheet you wanna create or a type "hssf" (.xsl) or "xssf" (.xlsx), by default "xssf".
	*/
	public static function newInstance(string pathOrType="xssf") {
		// check path
		var file="";
		var type="";
		if(!isNull(pathOrType) && len(pathOrType)>0) {
			if("xssf"==pathOrType) type="xssf";
			else if("xslx"==pathOrType) type="xssf";
			else if("hssf"==pathOrType) type="hssf";
			else if("xls"==pathOrType) type="hssf";
			else if("xlsx"==pathOrType) type="xssf";
			else {
				file=expandPath(pathOrType); 
				if(right(file,5)==".xlsx")  type="xssf";
				if(right(file,4)==".xls")  type="hssf";
			}
		}
		
		if(type=="") throw "could not find out type/path for [#pathOrType#]";

		if("xssf"==type) {
			local.proxy=loadPoi("org.apache.poi.xssf.usermodel.XSSFWorkbook");
		}
		if("hssf"==type) {
			local.proxy=loadPoi("org.apache.poi.hssf.usermodel.HSSFWorkbook");
		}
		return new POI(proxy.init(),file,true); 
	}


	/**
	* gets the date time format for dates written
	*/
	public string function getDateTimeFormat() localmode=true {
		return variables.dateTimeFormat;
	}
	/**
	* sets the date time format for dates written
	*/
	public any function setDateTimeFormat(required string format) localmode=true {
		variables.dateTimeFormat=arguments.format;
	}

	/**
	* gets the time format for dates written
	*/
	public string function getTimeFormat() localmode=true {
		return variables.timeFormat;
	}
	/**
	* sets the time format for dates written
	*/
	public any function setTimeFormat(required string format) localmode=true {
		variables.timeFormat=arguments.format;
	}

	/**
	* gets the date format for dates written
	*/
	public string function getDateFormat() localmode=true {
		return variables.dateFormat;
	}
	/**
	* sets the date format for dates written
	*/
	public any function setDateFormat(required string format) localmode=true {
		variables.dateFormat=arguments.format;
	}

	/**
	* returns all sheet names 
	*/
	public array function getSheetNames() localmode=true {
		arr=[];
		s=0;
		loop collection=variables.poi.sheetIterator() item="sheet" {
			arrayAppend(arr,sheet.getSheetName());
		}
		return arr;
	}

	/**
	* returns all sheet names 
	*/
	public string function getName(required numeric index) localmode=true {
		count=0;
		loop collection=variables.poi.sheetIterator() item="sheet" {
			count++;
			if(count==index) return sheet.getSheetName();
		}
		if(count==0) throw "there is no sheet with index [#index#], there are no sheets"; 
		if(count==1)throw "there is no sheet with index [#index#], valid index is only [1]"; 
		throw "there is no sheet with index [#index#], valid indexes are [1-#count#]"; 
	}

	/**
	* returns the index (1-n) of a sheet based on given name
	*/
	public numeric function getIndex(required string sheetName) localmode=true {
		index = variables.poi.getSheetIndex(sheetName);
		if(index==-1) throw "there is no sheet with name [#sheetName#]"; 
		return index+1;
	}

	/**
	* gets a specific cell (raw) 
	* @sheetName name of the sheet to write to, if sheet does not exists it will be created
	* @rowNbr number of row (1-n) to write to, if row does not exists it will be created
	* @colNbr number of column (1-n) to write to, if column does not exists it will be created
	* @createIfNecessary if set to true, the cell is created if it does not exists, if set to false it throws an exception in case the cell does not exists
	*/ 
	public any function getCell(required string sheetName, required numeric rowNbr, required numeric colNbr, 
		boolean createIfNecessary=true) localmode=true {
		sheet=_getSheet(sheetName,createIfNecessary);
		row=_getRow(sheet, rowNbr, createIfNecessary);
		cell=_getCell(row, colNbr, createIfNecessary,true);
		variables.isNew=false; // cell returned is possible manipulated
		return cell;
	}

	public void function copyRow(required string sheetName, required numeric sourceRowNum, 
		required numeric destinationRowNum, boolean copyContent=false) localmode=true {
        sheet=_getSheet(sheetName,true);

        // Get the source / new row
        newRow = sheet.getRow(destinationRowNum-1);
        sourceRow = sheet.getRow(sourceRowNum-1);

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (!isNull(newRow)) {
           	sheet.shiftRows(destinationRowNum-1, sheet.getLastRowNum(), 1);
        } 
        newRow = sheet.createRow(destinationRowNum-1);
        
        // Loop through source columns to add to new row
        for (var i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            oldCell = sourceRow.getCell(i);
            newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (isNull(oldCell)) {
                newCell = nullValue();
                continue;
            }

            // Copy style from old row and apply to new cell TODO optinize clone
            newRow.setRowStyle(sourceRow.getRowStyle());

            // Copy style from old cell and apply to new cell TODO optinize clone
            newCellStyle = variables.poi.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (!isNull(oldCell.getCellComment())) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (!isNull(oldCell.getHyperlink())) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            if(copyContent) {
	            switch (oldCell.getCellType()) {
	                case oldCell.CELL_TYPE_BLANK:
	                    newCell.setCellValue(oldCell.getStringCellValue());
	                    break;
	                case oldCell.CELL_TYPE_BOOLEAN:
	                    newCell.setCellValue(oldCell.getBooleanCellValue());
	                    break;
	                case oldCell.CELL_TYPE_ERROR:
	                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
	                    break;
	                case oldCell.CELL_TYPE_FORMULA:
	                    newCell.setCellFormula(oldCell.getCellFormula());
	                    break;
	                case oldCell.CELL_TYPE_NUMERIC:
	                    newCell.setCellValue(oldCell.getNumericCellValue());
	                    break;
	                case oldCell.CELL_TYPE_STRING:
	                    newCell.setCellValue(oldCell.getRichStringCellValue());
	                    break;
	            }
	        }
        }

        // If there are are any merged regions in the source row, copy to new row
        /*TODO for (var i = 0; i < sheet.getNumMergedRegions(); i++) {
            cellRangeAddress = sheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                        )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }*/
    }

	/**
	* write a style for a cell, row or sheet
	* @sheetName name of the sheet to write to, if sheet does not exists it will be created, if not defined style is applied to all cells in all sheets
	* @rowNbr number of row (1-n) to write to, this can be an integer, an array of integer or a closure, if row does not exists it will be created, if not defined style is applied to all rows in the sheet
	* @colNbr number of column (1-n) to write to, if column does not exists it will be created, if not defined style is applied to all columns in the sheet
	* @value can be a simple value, array, struct or query
	* @style style to apply
	*/ 
	public void function writeStyle(required string sheetName, any rowNbr=0, numeric colNbr=0, required struct style) localmode=true {
		sheet=_getSheet(sheetName,true);
		

		if(isCustomFunction(rowNbr)) {
			udf=rowNbr;
			rowNbr=[];
			lrn=sheet.getLastRowNum()+1;
			loop from="1" to=lrn index="rn" {
				if(udf(rn,lrn))	arrayAppend(rowNbr,rn);
			}
		}
		if(!isNull(rowNbr) && isSimpleValue(rowNbr) && rowNbr>0 ) {
			rowNbr=[rowNbr];
		}

		if(isArray(rowNbr)) {
			// TODO support arrays
			loop array=rowNbr item="rowNbr" {
				row=_getRow(sheet, rowNbr, true);
				// set a single cell
				if(colNbr>0) {
					cell=_getCell(row, colNbr, true);
					cell.setCellStyle(touchStyle(cell.getCellStyle(),style));	
					setColumnStyle(style, cell);
					setRowStyle(style, row);
				}
				// set all columns of that row
				else {
					row.setRowStyle(touchStyle(row.getRowStyle(),style));
					setRowStyle(style, row);
				}
			}
		}
		// all rows
		else {
			// single column
			if(colNbr>0) {
				sheet.setDefaultColumnStyle(colNbr-1,touchStyle(sheet.getColumnStyle(colNbr-1),style));
				setColumnStyle(style, sheet,colNbr-1);
				
			}
			else throw "you need to define a row or a column";
		}
		variables.isNew=false;
	}

	/**
	* sets the header to a sheet
	* @text text for header
	* @left text for the left part of the header, if set the text attribute is ignored
	* @center text for the center part of the header, if set the text attribute is ignored
	* @right text for the right part of the header, if set the text attribute is ignored
	*/
	public void function writeHeader(required string sheetName, string text='', string left='', string center='', string right='') localmode=true {
		var sheet=_getSheet(sheetName,true);
		var header=sheet.getHeader();
		loop list="text,left,center,right" item="local.type" {
			if(!isNull(arguments[type]) && len(arguments[type])>0) header["set"&type](arguments[type]);
		}
	}

	public struct function getHeader(required string sheetName) localmode=true {
		var sheet=_getSheet(sheetName,true);
		var header=sheet.getHeader();
		
		

		return {
			'left':header.getLeft(),
			'center':header.getCenter(),
			'right':header.getRight(),
			'text':extractText(header)
			//,'raw':header
		};
	}

	public struct function getFooter(required string sheetName) localmode=true {
		var sheet=_getSheet(sheetName,true);
		var footer=sheet.getFooter();

		return {
			'left':footer.getLeft(),
			'center':footer.getCenter(),
			'right':footer.getRight(),
			'text':extractText(footer)
			//,'raw':footer
		};
	}

	private function extractText(hf) localmode=true {
		var text="";
		try {
			var text=hf.getText();
		}
		catch(e) {
			try {
				var text=hf.getValue();
			}
			catch(e) {}
		}
		return text;
	}

	/**
	* sets the footer to a sheet
	* @text text for footer
	* @left text for the left part of the footer, if set the text attribute is ignored
	* @center text for the center part of the footer, if set the text attribute is ignored
	* @right text for the right part of the footer, if set the text attribute is ignored
	*/
	public void function writeFooter(required string sheetName, string text='', string left='', string center='', string right='') localmode=true {
		var sheet=_getSheet(sheetName,true);
		var footer=sheet.getFooter();
		loop list="text,left,center,right" item="local.type" {
			if(!isNull(arguments[type]) && len(arguments[type])>0) footer["set"&type](arguments[type]);
		}
	}

	/**
	* write a simple value, array, struct or query to a spreadsheet
	* @sheetName name of the sheet to write to, if sheet does not exists it will be created
	* @rowNbr number of row (1-n) to write to, if row does not exists it will be created
	* @colNbr number of column (1-n) to write to, if column does not exists it will be created
	* @value can be a simple value, array, struct or query
	* @type only used when you set a simple value, possible values are [auto,datetime,formula,number,boolean,string]
	* @addColumnNames only used when the value is a query, if set to true, the names of the columns of the query are also set into the spreadsheet
	* @horizontal fill the data horizontzally or vertically
	* @style  only used when you set a simple value, style used for the value
	* @comment only used when you set a simple value, set a comment for the value
	* @types types for the columns of a query, ignored for other types
	*/ 
	public void function write(required string sheetName, numeric rowNbr=1, numeric colNbr=1, 
		value, string type='auto', boolean addColumnNames=false, boolean horizontal=false,
		struct style, string comment='', types) localmode=true {
		
		//var start=getTickCount();

		sheet=_getSheet(sheetName,true);
		
		
		//systemOutput("- sheet:"&(getTickCount()-start),1,1);

		// is this a new, or d we already have set any values or styles?
		local.value=arguments.value;
		// query
		if(isQuery(value)) {
			
			
			//systemOutput("- get type:"&(getTickCount()-start),1,1);

			columnNames=queryColumnArray(value);
			
			// query columns
			if(addColumnNames) {
				if(!horizontal)row=_getRow(sheet, rowNbr, true);
				loop array=columnNames index="i" item="cn" {
					if(horizontal) {
						row=_getRow(sheet, rowNbr+i-1, true);
						cell=_getCell(row, colNbr, true);
					}
					else cell=_getCell(row, colNbr+i-1, true);
					writeValue(cell:cell,value:cn,type:'string',style:isNull(style)?nullValue():style);
				}
			}

		//systemOutput("- header:"&(getTickCount()-start),1,1);
			// query body
			iTimer = 0;
			var hasTypes=!isNull(arguments.types);
			
			var tmpTypes=[];
			if(hasTypes) {
				loop array=columnNames index="local.i" item="local.cn" {
					arrayAppend(tmpTypes,arguments.types[i]);
				}
			}
			else  {
				tmpTypes=extractTypes(value,false);
			}
			types=tmpTypes;	

			loop query=value {

				if(!horizontal) row=_getRow(sheet, rowNbr+value.currentrow-(addColumnNames?0:1), true);
				loop array=columnNames index="local.i" item="local.cn" {
					
					
					if(horizontal) {
						row=_getRow(sheet, rowNbr+i-1, true);
						cell=_getCell(row, colNbr+value.currentrow-(addColumnNames?0:1), true);
					}
					else {
						cell=_getCell(row, colNbr+i-1, true);
					}
					
					var val=QueryGetCell(value,cn,QueryCurrentRow(value));
					if(isStruct(val) && structKeyExists(val,"value") && structKeyExists(val,"style")) {
						var style=val.style?:nullValue();
						var val=val.value?:nullValue();
					}
					
					if( (types[i]?:"auto")!="auto" && !isExtendedValue(val)) {
						writeValue(
							cell:cell,
							value:val,
							style:isNull(style)?nullValue():style,
							type:types[i]
						);

					
					}
					else {	
						writeValueAuto(
							cell, 
							val, 
							isNull(style)?nullValue():style
						);

					}
				}

			}

		//systemOutput("- body:"&(getTickCount()-start),1,1);
		}
		// array
		else if(isArray(value)) {
			if(horizontal)row=_getRow(sheet, rowNbr, true);
			loop array=value index="i" item="v" {
				if(horizontal) cell=_getCell(row, colNbr+i-1, true);
				else {
					row=_getRow(sheet, rowNbr+i-1, true);
					cell=_getCell(row, colNbr, true);
				}
				writeValueAuto(cell,v,isNull(style)?nullValue():style);
			}
		}
		// struct
		else if(isStruct(value)) {
			// a struct can also be a simple value with addional info to the value
			if(isExtendedValue(value)) {
				row=_getRow(sheet, rowNbr, true);
				cell=_getCell(row, colNbr, true);
				writeValueAuto(cell,value,isNull(style)?nullValue():style);
			}
			else {
				if(horizontal){
					rowK=_getRow(sheet, rowNbr, true);
					rowV=_getRow(sheet, rowNbr+1, true);
				}
				i=0;
				loop struct=value index="k" item="v" {
					i++;
					if(horizontal) {
						cellK=_getCell(rowK, colNbr+i-1, true);
						cellV=_getCell(rowV, colNbr+i-1, true);
					}
					else {
						row=_getRow(sheet, rowNbr+i-1, true);
						cellK=_getCell(row, colNbr, true);
						cellV=_getCell(row, colNbr+1, true);
					}
					writeValue(cell:cellK,value:k,type:'string',style:isNull(style)?nullValue():style);
					writeValueAuto(cellV,v,isNull(style)?nullValue():style);
				}
			}
		}
		// other
		else { // check null for value
			row=_getRow(sheet, rowNbr, true);
			cell=_getCell(row, colNbr, true);
			writeValue(cell,value,type,isNull(style)?nullValue():style,isNull(comment)?nullValue():comment);
		}
		variables.isNew=false;
	}

	/**
	* store spreadsheet to a file
	* @path location of the spreadsheet to write, optional if a existing spreedsheet was modified
	*/ 
	public void function store(path='') localmode=true {
		if(len(arguments.path)==0) {
			if(len(variables.file)==0) throw "missing path definition";
			else local.file=variables.file;
		}
		else local.file=path;


		if(right(file,5)==".xlsx" && "xssf"!=getType()) throw "this spreadsheet cannot be stored as xssf (.xlsx) document.";
		if(right(file,4)==".xls" && "hssf"!=getType())  throw "this spreadsheet cannot be stored as hssf (.xls) document.";
		// TODO convert from hssf to xssf (visa versa) http://stackoverflow.com/questions/7230819/how-to-convert-hssfworkbook-to-xssfworkbook-using-apache-poi

		// create file if necessary
		if(!fileExists(file)) {
			fileWrite(file,'');
		}
		FileOutputStream=POI::loadSystem("java.io.FileOutputStream");
		//_File=loadSystem("java.io.File");
		os=FileOutputStream.init(file);
		try{
			poi.write(os);
		}
		finally {
			os.close();
		}
	}

	private function _getSheet(required string sheetName,required boolean createIfNecessary) localmode=true {
		index =  variables.poi.getNumberOfSheets()==0? -1 : variables.poi.getSheetIndex(sheetName);
		if(index==-1 ) {
			if(!createIfNecessary) throw "there is no sheet with name [#sheetName#]"; 
			return variables.poi.createSheet(sheetName);
		}
		return variables.poi.getSheetAt(index);
	}

	private function _getRow(required sheet, required numeric rowNbr,required boolean createIfNecessary) localmode=true {
		rn=rowNbr-1;
		if(rn<0) throw "invalid row number [#rowNbr#]"
		
		row=sheet.getRow(rn);
		if(isNull(row)) {
			if(!createIfNecessary && sheet.getLastRowNum()<rn) throw "row #rowNbr# does not exist"
			row=sheet.createRow(rn);
		}
		return row;
	}

	private function _getCell(required row, required numeric colNbr,required boolean createIfNecessary) localmode=true {
		cn=colNbr-1;
		if(cn<0) throw "invalid cell number [#colNbr#]"
		if(!variables.isNew) cell=row.getCell(cn);
		if(variables.isNew || isNull(cell)) {
			//systemOutput("create cell "& variables.isNew,1,1);
			if(!createIfNecessary && row.getLastCellNum()<cn) throw "cell #colNbr# does not exist"
			cell=row.createCell(cn);

			if(!variables.isNew) {
				// set style
				rowStyle=row.getRowStyle();
				colStyle=row.getSheet().getColumnStyle(cn);
				
				if(!isNull(rowStyle) || !isNull(colStyle)) {
					key=(isNull(rowStyle)?'-1':rowStyle.getIndex())&":" & (isNull(colStyle)?'-1':colStyle.getIndex());
					
					if(structKeyExists(variables.rowColStyles,key)) {
						tmp=variables.rowColStyles[key];
					}
					else {
						createHelper = variables.poi.getCreationHelper();
						tmp = variables.poi.createCellStyle();
						//systemOutput("createStyle 411:"&key ,1,1);
						if(!isNull(rowStyle))tmp.cloneStyleFrom(rowStyle); 
						if(!isNull(colStyle))tmp.cloneStyleFrom(colStyle); 
						variables.rowColStyles[key]=tmp;
					}
					cell.setCellStyle(tmp);
				}
			}
		} 
		return cell;
	}

	private void function writeValueAuto(required cell,required value, struct style) localmode=true {
		if(isExtendedValue(arguments.value)) {
			writeValueExtended(cell,value,isNull(arguments.style)?nullValue():arguments.style);
		}
		else {
			writeValue(cell:arguments.cell,value:arguments.value,style:isNull(arguments.style)?nullValue():arguments.style);
		}
	}

	private void function writeValueExtended(required cell,required value, struct style) localmode=true {
		if(!isNull(arguments.style)) _style=arguments.style;
		else if(!isNull(value.style)) _style=value.style;

		writeValue(cell:cell,
			value:isNull(value.value)?nullValue():value.value,
			type:isNull(value.type)?'auto':value.type,
			style:isNull(_style)?nullValue():_style,
			comment:isNull(value.comment)?nullValue():value.comment);
	}

	private void function writeValue(required cell,required value, string type='auto', struct style, comment) localmode=true {
		local.value=arguments.value;
		//systemoutput(type,1,1);
		if(type=='auto') {
			
			// first of all we need to look for a type from row or cell
			defType=getCellType(cell);
			
			if(defType=="datetime" || defType=="numeric" || defType=="string" || defType=="formula" || defType=="boolean")
				type=defType;
			// order here is important, because isDate will aceept number like 12.12 as a date, so before we check for date we have to make sure it is not a number
			else if(isNumeric(value)) type='number';
			else if(!StringUtils.isAlphaSpace(value) && isDate(value)) { // workaround because isDate is very slow if "isAlphaSpace" is true we have no number, so it cannot be a number
				tmp=lsparseDateTime(value);
				type='datetime';
				if (year(tmp) lt 1900 || year(tmp) gt 9999 ) type='string';
				else value=tmp;
			}
			else if(isBoolean(value)) type='boolean';
			else type='string';
			//systemOutput(type,1,1);
		}

		// value
		if(!isNull(value)) {
			try {
				if(type=='string' || value=="") {
					//variables.STCellType=createObject('java','org.xlsx4j.sml.STCellType');
					//var ct=cell.getCTCell();
					//ct.setT(STCellType.INLINE_STR);

					cell.setCellValue(javaCast('string',value));
				}
				else if(type=='number' || type=='numeric') {
					value = replace(value, ",", ".", "ONCE");
					if(value=="-" || value=="+") { // explizite nicht numerische character die als Zahl akzeptiert werden
						cell.setCellValue(value);
					}
					else {
						cell.setCellValue(javaCast('double',val(value)));
					}
				}
				else if(type=='boolean') cell.setCellValue(javaCast('boolean',value));
				else if(type=='formula') cell.setCellFormula(javaCast('string',a.value));
				else if(type=='datetime' || type=='date') {
					
					dt=lsparseDateTime(value);
					if(hour(dt)==0 && minute(dt)==0 && minute(dt)==0 && second(dt)==0 && millisecond(dt)==0)
						format=isNull(style.dateFormat)?variables.dateFormat:style.dateFormat;
					//else if(year(dt)==1899 && month(dt)==12 && day(dt)==30)
					//	format=variables.timeFormat;
					else format=isNull(style.dateTimeFormat)?variables.dateTimeFormat:style.dateTimeFormat;

					var cellStyle=cell.getCellStyle();
					var key=cellStyle.getIndex()&":"&hash(format);
					if(!structKeyExists(variables.dateStyles,key)) {
						createHelper = variables.poi.getCreationHelper();
						//systemOutput("createStyle 483:" & key,1,1);
						tmp = variables.poi.createCellStyle();
						tmp.cloneStyleFrom(cellStyle); 
						tmp.setDataFormat(createHelper.createDataFormat().getFormat(format));
						local.cellStyle=variables.dateStyles[key]=tmp;
					}
					else local.cellStyle=variables.dateStyles[key];
					
					cell.setCellValue(javaCast('java.util.Date',dt));
					cell.setCellStyle(cellStyle);
				    // TODO should we not simply do  getCellStyle 
				}
				else throw "type [#type#] is not supported, supported types are [auto,datetime,boolean,formula,number,string]";
			}
			catch(e) {
				// systemOutput("->"&value,1,1);
				cell.setCellValue(javaCast('string',value));
			}
		}

		// style 
		if(!isNull(style) && structCount(style)>0) {
			var cs=touchStyle(cell.getCellStyle(),style);
			cell.setCellStyle(cs);	
			setColumnStyle(style, cell);
			setRowStyle(style, cell.getRow());
		}
		else {
			//dump(cell.getCellStyle());
		}

		// comment TODO
		if(!isNull(comment) && len(comment)>0) {
			cell.setCellComment(toComment(cell,comment));
		}
	}

	private string function createKey(required struct data) {
		return hash(data.toString());
	}

	private function touchStyle(style,struct data) {
		//var start=getTickCount('nano');
		//if(isNull(cs) && !isNull(cell)) cs=cell.getCellStyle();
		
		var key=(isNull(style)?-1:style.getIndex())&":"&createKey(data);
		if(!structKeyExists(variables.cellStyles,key)) {
			createHelper = variables.poi.getCreationHelper();
			//systemOutput("createStyle 526",1,1);
			local.cellStyle = variables.poi.createCellStyle();
			if(!isNull(style))local.cellStyle.cloneStyleFrom(style);
			setStyle(local.cellStyle,data); 
			variables.cellStyles[key]=local.cellStyle;
		}
		else local.cellStyle=variables.cellStyles[key];
		//dump(getTickCount('nano')-start);
		return local.cellStyle;
	}
	

	/**
	* reads a sheet defined by name
	* @name name of the sheet to read
	* @extended if true the value of a single cell is a struct containing all kind of info to that cell, if false only the value itself is provided
	* @suppressEmptyColumnsBefore if set to true empty columns before existing data are supressed
	* @suppressEmptyRowsBefore if set to true empty rows before existing data are supressed
	* @firstLineAsColumnNames first line of the data read are used as column names of the returned query
	* @debug if set to true dump out debug information
	*/ 
	public any function readSheetByName(required string name, boolean extended=false, 
		boolean suppressEmptyColumnsBefore=false, boolean suppressEmptyRowsBefore=false,
		boolean firstLineAsColumnNames=false, debug=false) localmode=true {
		
		return readSheetByIndex(variables.poi.getSheetIndex(name)+1, extended,suppressEmptyColumnsBefore, suppressEmptyRowsBefore, 
			firstLineAsColumnNames, debug);
	}

	/**
	* reads a sheet defined by name
	* @index index of the sheet to read
	* @extended if true the value of a single cell is a struct containing all kind of info to that cell, if false only the value itself is provided
	* @suppressEmptyColumnsBefore if set to true empty columns before existing data are supressed
	* @suppressEmptyRowsBefore if set to true empty rows before existing data are supressed
	* @firstLineAsColumnNames first line of the data read are used as column names of the returned query
	* @debug if set to true dump out debug information
	*/ 
	public any function readSheetByIndex(required numeric index, boolean extended=false, 
		boolean suppressEmptyColumnsBefore=false, boolean suppressEmptyRowsBefore=false,
		boolean firstLineAsColumnNames=false, debug=false) localmode=true {
		
		try {
			sheet=variables.poi.getSheetAt(index-1);

		}
		catch(java.lang.IllegalArgumentException e) {
			throw "there is no sheet with index [#index#]";
		}
		variables.isNew=false;
		return readSheet(sheet,extended,suppressEmptyColumnsBefore,suppressEmptyRowsBefore, 
			firstLineAsColumnNames, debug);
	}

	

	/**
	* reads all sheets defined in the spreadsheet
	* @extended if true the value of a single cell is a struct containing all kind of info to that cell, if false only the value itself is provided
	* @suppressEmptyColumnsBefore if set to true empty columns before existing data are supressed
	* @suppressEmptyRowsBefore if set to true empty rows before existing data are supressed
	* @firstLineAsColumnNames first line of the data read are used as column names of the returned query
	* @debug if set to true dump out debug information
	*/ 
	public any function readSheets(boolean extended=false, 
		boolean suppressEmptyColumnsBefore=false, boolean suppressEmptyRowsBefore=false,
		boolean firstLineAsColumnNames=false, boolean debug=false) localmode=true {
		arr=[];
		loop collection=variables.poi.sheetIterator() item="sheet" {
			arrayAppend(arr,readSheet(sheet,extended,suppressEmptyColumnsBefore,suppressEmptyRowsBefore, 
			firstLineAsColumnNames,debug));
		}
		if(!extended) return arr;
		
		sct=structNew("linked");
		sct['sheets']=arr;
		sct['version']=getSpreadsheetVersion();
		sct['NumCellStyles']=variables.poi.getNumCellStyles();

		variables.isNew=false;
		return sct;
	}

	public any function info() localmode=true {
		
		sct=structNew("linked");
		sct['version']=getSpreadsheetVersion();
		sct['NumberOf']['CellStyles']=variables.poi.getNumCellStyles();
		sct['NumberOf']['Sheets']=variables.poi.getNumberOfSheets();
		sct['NumberOf']['Fonts']=variables.poi.getNumberOfFonts();
		sct['NumberOf']['Names']=variables.poi.getNumberOfNames();

		// info to sheets
		sct['sheets']=structNew("linked");
		loop collection=variables.poi.sheetIterator() item="sheet" {
			var ssct=structNew("linked");
			var frn=sheet.getFirstRowNum();
			var lrn=sheet.getLastRowNum();
			if(frn>=0) local.frow=sheet.getRow(frn);
			if(lrn>=0) local.lrow=sheet.getRow(lrn);
			ssct['firstRowNumber']=frn+1;
			ssct['lastRowNumber']=lrn+1;
			if(!isnull(frow))ssct['firstCellNumberFromFirstRow']=frow.getFirstCellNum()+1;
			if(!isnull(lrow))ssct['firstCellNumberFromLastRow']=lrow.getFirstCellNum()+1;
			if(!isnull(frow)) {
				var l=frow.getLastCellNum();
				ssct['lastCellNumberFromFirstRow']=l<0?0:l;
			}
			if(!isnull(lrow)) {
				var l=lrow.getLastCellNum();
				ssct['lastCellNumberFromLastRow']=l<0?0:l;
			}
			sct['sheets'][sheet.getSheetName()]=ssct;
		}


		return sct;
	}

	private any function readSheet(required sheet,boolean extended=false, 
		boolean suppressEmptyColumnsBefore=false, boolean suppressEmptyRowsBefore=true,
		boolean firstLineAsColumnNames=false, boolean debug=false) localmode=true {
		sheetName=sheet.getSheetName();
		sheetIndex=sheet.getWorkbook().getSheetIndex(sheetName)+1;
		qry=queryNew("");
		if(extended) {
			_data=structNew("linked");
			_data['Autobreaks']=sheet.getAutobreaks();
			_data['DefaultColumnWidth']=sheet.getDefaultColumnWidth();
			_data['DefaultRowHeight']=sheet.getDefaultRowHeight();
			_data['DefaultRowHeightInPoints']=sheet.getDefaultRowHeightInPoints();
			_data['DisplayGuts']=sheet.getDisplayGuts();
			_data['FirstRowNumber']=sheet.getFirstRowNum()+1;
			_data['FitToPage']=sheet.getFitToPage();
			_data['ForceFormulaRecalculation']=sheet.getForceFormulaRecalculation();
			_data['HorizontallyCenter']=sheet.getHorizontallyCenter();
			_data['LastRowNum']=sheet.getLastRowNum();
			_data['LeftCol']=sheet.getLeftCol();
			_data['PhysicalNumberOfRows']=sheet.getPhysicalNumberOfRows();
			_data['Protect']=sheet.getProtect();
			_data['TopRow']=sheet.getTopRow();
			_data['VerticallyCenter']=sheet.getVerticallyCenter();
			_data['DisplayFormulas']=sheet.isDisplayFormulas();
			_data['DisplayGridlines']=sheet.isDisplayGridlines();
			_data['DisplayRowColHeadings']=sheet.isDisplayRowColHeadings();
			_data['DisplayZeros']=sheet.isDisplayZeros();
			_data['PrintGridlines']=sheet.isPrintGridlines();
			_data['RightToLeft']=sheet.isRightToLeft();
			_data['Selected']=sheet.isSelected();
			_data['name']=sheet.getSheetName();
			_data['index']=variables.poi.getSheetIndex(sheet)+1;
			_data['data']=qry;

			// HSS Specific
			if("hssf"==getType()) {
				_data['AlternateExpression']=sheet.getAlternateExpression();
				_data['AlternateFormula']=sheet.getAlternateFormula();
				_data['Dialog']=sheet.getDialog();
				_data['Active']=sheet.isActive();
				_data['GridsPrinted']=sheet.isGridsPrinted();
			}
			// XSS Specific
			if("xssf"==getType()) {
				_data['WindowsLocked']=variables.poi.isWindowsLocked();
				
			}
			// TODO add XSS specifc values -> https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/XSSFWorkbook.html
			
		} 
		else _data=qry;
		//data['header']=sheet.getHeader(); // TODO convert in CF  types
		//data['foother']=sheet.getFooter(); // TODO convert in CF  types
		defaultColumnWidth=sheet.getDefaultColumnWidth();
		columnWidths=structNew("linked");
		columnStyles=structNew("linked");
		rowStyles=structNew("linked");

		doneBeforeColumns=false;
		doneCell=false;
		lastRN=0;
		rowCount=0;
		loop collection=sheet.rowIterator() item="row" {
			rowCount++;
			rn=row.getRowNum()+1;
			r=queryAddRow(qry,suppressEmptyRowsBefore && !doneCell?1:rn-lastRN);
			fcn=row.getFirstCellNum()+1;
			lastRN=rn;
			// add empty columns before the data
			if(!doneBeforeColumns && !suppressEmptyColumnsBefore && fcn>1) {
				for(i=1;i<fcn;i++){
					cn=getColumnName(i);
					if(!queryColumnExists(qry,cn)) {
						queryAddColumn(qry, cn,[]);
						doneBeforeColumns=true;
					}
				}
			}
			if(arguments.debug && sheetIndex==1 && r==1)dump(row);


			if(isNull(rowStyles[r])) {
				rowStyles[r]=getStyle(row.getRowStyle());

				if(isNull(rowStyles[r]))rowStyles[r]={};
				rowStyles[r]['HeightInPoints']=row.getHeightInPoints();
				rowStyles[r]['Height']=row.getHeight();
			}
			//if(r==10) break;
			c=0;
			loop collection=row.cellIterator() item="cell" {
				c++;
				ci=cell.getColumnIndex()+1;
				colName=getColumnName(ci);
				if(isNull(columnWidths[ci])) {
					try {
						columnWidths[ci]={
							'ColumnWidth':sheet.getColumnWidth(ci-1),
							'ColumnWidthInPixels':sheet.getColumnWidthInPixels(ci-1)
						};
					} catch (local.e) {
					}
				}
				if(isNull(columnStyles[ci])) {
					columnStyles[colName]=getStyle(sheet.getColumnStyle(ci));
					if(isNull(columnStyles[colName]))columnStyles[colName]={};
					columnStyles[colName]['WidthInPixels']=columnWidths[ci].ColumnWidthInPixels;
					columnStyles[colName]['Width']=columnWidths[ci].ColumnWidth;
				}

				cn="c"&(ci);
				cn=colName;
				if(!queryColumnExists(qry,cn)) { // first row
					queryAddColumn(qry, cn,[]);

				}
				if(!extended) {
					value=getValue(cell); 
				}
				else {
					value=structNew("linked");
					value['columnIndex']=ci;
					value['columnName']=colName;
					value['rowNumber']=rn;
					value['type']=getCellType(cell);
					value['value']=getValue(cell);
					value['comment']=getComment(cell);
					value['style']=getStyle(cell.getCellStyle());
					value['style']['WidthInPixels']=columnWidths[ci].ColumnWidthInPixels;
					value['style']['Width']=columnWidths[ci].ColumnWidth;
					value['style']['Height']=row.getHeight();
					
					//value['address']=cell.getAddress();

				}
				doneCell=true;
				querySetCell(qry, cn, value, r);
				
				if(debug && sheetIndex==1 && r==1 && c==1)dump(cell);
				
			}

		}

		if(extended) {
			_data['column']['style']=columnStyles;
			_data['row']['style']=rowStyles;
		}

		// convert column names
		if(firstLineAsColumnNames && qry.recordcount) {
			qry2=queryNew('');
			loop array=queryColumnArray(qry) item="col" {
				data=queryColumnData(qry,col);
				first=qry[col][1];
				if(isStruct(first)) first=first.value;
				if(len(first)==0) first=col;
				arrayDeleteAt(data,1);

				newColName=first;
				cnt=0;
				while(queryColumnExists(qry2,newColName)) newColName=first&" ("&(++cnt)&")";
				queryAddColumn(qry2,newColName,data);
			}
			if(extended)_data.data=qry2;
			else _data=qry2;
		}


		variables.isNew=false;
		return _data;
		
	}


	/**
	* returns the raw data object
	*/ 
	public any function getRaw() localmode=true {
		variables.isNew=false;
		return variables.poi;
	}

	public string function getType() localmode=true {
		if(isNull(variables.type)) {
			className=listLast(variables.poi.getClass().getName(),'.');
			if("HSSFWorkbook"==className) variables.type = "hssf";
			else if("XSSFWorkbook"==className) variables.type = "xssf";
			else variables.type = replace(className,'Workbook','');
		}
		return variables.type;
	}


	variables.letters=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
	private string function getColumnName(required numeric index) localmode=true {
		// TODO only support up to 26*26
		len=variables.letters.len()
		last=index%len;
		first=int(index/len);
		if(last==0) {
			last=len;
			first--;
		}
		if(first>0) return variables.letters[first]&variables.letters[last]
		return variables.letters[last];
		
	}

	private string function toColorShort(required string color) cachedwithin="request" localmode=true {
		// return -1;
		if("auto"==color) return -1;

		c=variables.ColorCaster.toColor(color); // TODO support for ARGB color? 
		
		if("xssf"==getType()) {
			return XSSFColor.init(c);
		}
		
		palette=variables.poi.getCustomPalette();
		
		// do we have a match?
		finding=palette.findColor(c.getRed(),c.getGreen(),c.getBlue());
		//if(isNull(finding))finding=palette.findSimilarColor(c.getRed(),c.getGreen(),c.getBlue());
		
		// do we have a match 
		if(!isNull(finding)) {
			triplet=finding.getTriplet();
			if(c.getRed()==triplet[1] && c.getGreen()==triplet[2] && c.getBlue()==triplet[3]) {
				return  (finding).getIndex();
			}
		}

		// create a new custom color if we have no match
		param request.sPalettesAdded = 0;
		try {

			request.sPalettesAdded++;
			return  (palette.addColor(c.getRed(), c.getGreen(), c.getBlue())).getIndex() ;
		}
		catch(java.lang.RuntimeException re) {
			if(re.message=="Could not find free color index") {
				var freeIndex=getColorIndexNotUsed(palette);
				if(freeIndex>0) {
					palette.setColorAtIndex(freeIndex,c.getRed(), c.getGreen(), c.getBlue());
					return freeIndex;
				}
			}
			rethrow;	
		}
		
	}

	public function getColors() {
		var indexes=getUsedColorIndexes(true);
		var q=queryNew("index,hex");
		loop struct=indexes index="local.index" item="local.hex" {
			if(hex=="auto") continue;
			var row=queryAddRow(q);
			querySetCell(q,"index",index,row);
			querySetCell(q,"hex",hex,row);
		}
		return q;
	}


	private string function getColorIndexNotUsed(palette) localmode=true {
		var indexes=getUsedColorIndexes();

		loop from=1 to=1000 index="local.i" {
			if(!structKeyExists(indexes,i)) {
				var c=palette.getColor(i);
				if(!isNull(c)) {
					return i;
				}
			}
		}
		return 0;
	}


	private array function getUsedColorIndexes(values=false) localmode=true {
		var indexes={};
		var it=variables.poi.sheetIterator();
		while(it.hasNext()) {
			var sheet=it.next();
			var itt=sheet.rowIterator();
			
			// row
			while(itt.hasNext()) {
				var row=itt.next();
				getIndexesFromStyle(indexes,row.getRowStyle(),values);
				var ittt=row.cellIterator();
				while(ittt.hasNext()) {
					var cell=ittt.next();
					getIndexesFromStyle(indexes,cell.getCellStyle(),values);
				}
			}

			// column
			var col=-1;
			var last=-1;
			while(col++<1000) {
				var style=sheet.getColumnStyle(col);
				if(!isNull(style)){
					getIndexesFromStyle(indexes,style,values);
					last=col;
				}
				if(last+25<col) break;
			}
		}
		structDelete(indexes,"0");
		return indexes;
	}

	private string function getIndexesFromStyle(indexes,style,values=false) localmode=true {
		if(isNull(style)) return;
		indexes[style.getFillBackgroundColor()]=values?toColor(style.getFillBackgroundColorColor()):'';
		indexes[style.getFillForegroundColor()]=values?toColor(style.getFillForegroundColorColor()):'';
		if("xssf"==getType()) {
			indexes[style.getLeftBorderColor()]=values?toColor(style.getLeftBorderXSSFColor()):'';
			indexes[style.getTopBorderColor()]=values?toColor(style.getTopBorderXSSFColor()):'';
			indexes[style.getRightBorderColor()]=values?toColor(style.getRightBorderXSSFColor()):'';
			indexes[style.getBottomBorderColor()]=values?toColor(style.getBottomBorderXSSFColor()):'';

			var font=style.getFont();
			indexes[font.getColor()]=values?toColor(font.getXSSFColor()):'';
		}
		else {
			if(values) local.palette=variables.poi.getCustomPalette();
		
			indexes[style.getLeftBorderColor()]=values?toColor(palette.getColor(style.getLeftBorderColor())):'';
			indexes[style.getTopBorderColor()]=values?toColor(palette.getColor(style.getTopBorderColor())):'';
			indexes[style.getRightBorderColor()]=values?toColor(palette.getColor(style.getRightBorderColor())):'';
			indexes[style.getBottomBorderColor()]=values?toColor(palette.getColor(style.getBottomBorderColor())):'';

			var font=style.getFont(variables.poi);
			indexes[font.getColor()]=values?toColor(palette.getColor(font.getColor())):'';
		}
		

	}


	
	private string function toColor(color) localmode=true {
		
		if(isNull(color)) {
			return "auto";
		}
		if(isNumeric(color)) {
			if("xssf"==getType()) {
				throw "xssf does not support index colors!";
			}
			c=poi.getCustomPalette().getColor(color);
			
			if(isNull(c)) {
				return 'auto';// TODO getParent Color
			}
			else return toColor(c);
		}
		if("org.apache.poi.xssf.usermodel.XSSFColor"==color.getClass().getName()) {
			
			key="color:"&color.getARGBHex();
			//if(structKeyExists(variables.colors,key)) return variables.colors[key];

			if(color.isAuto()) return "auto";
			argb=color.getARGB();
			if(isNull(argb)) return ""; // TODO handle this situation, seems not to be auto
			

			var hex=color.getARGBHex();
			if(len(hex)==8 && left(hex,2)=="ff") {
				hex=mid(hex,3);
			}
			return "##"&hex;
			

			//if(argb[1]>-1) return variables.colors[key]="##"&toHex(round(argb[1]*255/100))&toHex(argb[2])&toHex(argb[3])&toHex(argb[4]);
			//return variables.colors[key]="##"&toHex(argb[2])&toHex(argb[3])&toHex(argb[4]);
		}

		if(!findNoCase("custom",color&"")) {
			return "auto";
		}
		arr=color.getTriplet();
		return "##"&toHex(arr[1],16)&toHex(arr[2],16)&toHex(arr[3],16);
	}

	private string function toHex(required numeric nbr) localmode=true {
		if(nbr==-1) return "ff";
		res = formatBaseN(nbr,16);
		if(res.len()==1) return "0"&res;
		return res;
	}

	variables.DateUtil=POI::loadPoi("org.apache.poi.ss.usermodel.DateUtil");
	private string function getCellType(required cell) localmode=true {
		
		switch(cell.getCellType()) {
			case cell.CELL_TYPE_BLANK: return "blank";
			case cell.CELL_TYPE_STRING: return "string";
			case cell.CELL_TYPE_NUMERIC: 
				if(DateUtil.isCellDateFormatted(cell)) return "datetime";
				return "numeric";
			case cell.CELL_TYPE_FORMULA: return "formula";
			case cell.CELL_TYPE_BOOLEAN: return "boolean";
			case cell.CELL_TYPE_ERROR: return "error";
		}

		return "unknown";
	}

	private function toComment(cell, any data) localmode=true {
		if(isSimpleValue(data)) data={value:data};


		factory = variables.poi.getCreationHelper();
		drawing =cell.getSheet().createDrawingPatriarch();
		row=cell.getRow();
		// TODO make this better
		anchor = factory.createClientAnchor();
	    anchor.setCol1(cell.getColumnIndex());
	    anchor.setCol2(cell.getColumnIndex()+1);
	    anchor.setRow1(row.getRowNum());
	    anchor.setRow2(row.getRowNum()+3);
	    
	    // Create the comment and set the text+author
	    comment = drawing.createCellComment(anchor);
	    
	    // value
	    if(!isNull(data.value)) {
	    	str = factory.createRichTextString(data.value);
	    	comment.setString(str);
	    }
	    
	    // author
	    if(!isNull(data.author)) comment.setAuthor(data.author);

	    // TODO hssf should support more

	    return comment;
	}

	private any function getComment(required cell) localmode=true {
		c=cell.getCellComment();
		if(isNull(C)) return "";

		sct=structNew("linked");
		sct['author']=c.getAuthor();
		sct['visible']=c.isVisible();

		sct['value']=getString(c.getString());
		sct['author']=c.getAuthor();

		if("hssf"==getType()){
			sct['HorizontalAlignment']=getHorizontalAlignment(c);
			sct['VerticalAlignment']=getVerticalAlignment(c);
			sct['margin']['left']=c.getMarginLeft();
			sct['margin']['top']=c.getMarginTop();
			sct['margin']['right']=c.getMarginRight();
			sct['margin']['bottom']=c.getMarginBottom();
			sct['Flip']['Horizontal']=c.isFlipHorizontal();
			sct['Flip']['Vertical']=c.isFlipVertical();
			sct['RotationDegree']=c.getRotationDegree();
			sct['CountOfAllChildren']=c.countOfAllChildren();
			sct['NoFill']=c.isNoFill();
			sct['FillColor']=toColor(c.getFillColor());
			sct['LineStyleColor']=toColor(c.getLineStyleColor());
			sct['LineWidth']=c.getLineWidth();
			sct['LineStyle']=getLineStyle(c);
		}
		// TODO XSSF

		// TODO
		// getBackgroundImageId()  
		// getClientAnchor()
		// getNoteRecord() 
		// getOptRecord
		// getAnchor()




		return sct;
	}

	private function setStyle(required any style,required struct data) localmode=true {
		
		// alignment
		if(!isNull(data.Alignment))  		style.setAlignment(toHorizontalAlignment(data.Alignment));
		if(!isNull(data.verticalAlignment))	style.setVerticalAlignment(toVerticalAlignment(data.verticalAlignment));

		
		// TODO what is the diff between BG and FG
			
		// Foreground Color
		if((hasFill=(!isNull(data.FillForegroundColor))) || !isNull(data.ForegroundColor)) { 
			fg=hasFill?data.FillForegroundColor:data.ForegroundColor;
			style.setFillForegroundColor(toColorShort(fg));
			if(isNull(data.FillPattern)) data.FillPattern="SOLID_FOREGROUND";
		}
		// Background Color
		else if((hasFill=(!isNull(data.FillBackgroundColor))) || !isNull(data.BackgroundColor)) {
			bg=hasFill?data.FillBackgroundColor:data.BackgroundColor;
			style.setFillForegroundColor(toColorShort(bg));
			if(isNull(data.FillPattern)) data.FillPattern="SOLID_FOREGROUND";
		}

		// Border
		if(!isNull(data.Border)) setBorder(style,data.Border);
		// Border color
		if(!isNull(data.BorderColor)) {
			setBorderColor(style,data.BorderColor);
		}
		// Fill Pattern (do not change order of this)
		if(!isNull(data.FillPattern)) style.setFillPattern(toFillPattern(data.FillPattern));
		// Font
		if(!isNull(data.Font)) setFont(style,data.Font);
		// Hidden
		if(!isNull(data.Hidden)) style.setHidden(data.Hidden==true);
		// Indention
		if(!isNull(data.Indention)) style.setIndention(int(data.Indention));
		// Locked 
		if(!isNull(data.Locked)) style.setLocked(data.Locked==true);
		// QuotePrefixed 
		if(!isNull(data.QuotePrefixed)) style.setQuotePrefixed(data.QuotePrefixed==true);
		// ReadingOrder
		
		// ShrinkToFit
		if(!isNull(data.ShrinkToFit)) style.setShrinkToFit(data.ShrinkToFit==true);

		//if(!isNull(data.))  sct['DataFormat']=style.getDataFormatString(poi);
		

		//if(!isNull(data.))  sct['UserStyleName']=style.getUserStyleName();
		//if(!isNull(data.))  sct['VerticalAlignmentEnum']=data.getVerticalAlignmentEnum().toString();

		if(!isNull(data.WrapText)) {
//			if (debug) DUMP(data.WrapText);
			style.setWrapText(data.WrapText==true);
		}
		if("hssf"==getType()) {
			if(!isNull(data.Rotation)) style.setRotation(int(data.Rotation));
			if(!isNull(data.ReadingOrder)) style.setReadingOrder(toReadingOrder(data.ReadingOrder));
		}
		if("xssf"==getType()) {
			// Rotation TODO if XSS -90 is this correct?
			if(!isNull(data.Rotation)) style.setRotation(int(data.Rotation));
		}
		// TODO XSS

	}

	private function setColumnStyle(required struct data, any cellOrSheet, numeric index=-1) localmode=true {
		if(isNull(cellOrSheet)) return;

		isSheet=findNoCase("Sheet",cellOrSheet.getClass().getName())>0;
		// column with 
		if(!isSheet) {
			cell=cellOrSheet;
			index= cell.getColumnIndex();
			sheet=cell.getSheet();
		}
		else {
			sheet=cellOrSheet;
		}

		if(!isNull(data.WidthInPixels)) {
			data.WidthInPixels = len(data.WidthInPixels)==0 ? 100 : data.WidthInPixels;
			ratio=sheet.getColumnWidth(index)/sheet.getColumnWidthInPixels(index);
			sheet.setColumnWidth(index, data.WidthInPixels*ratio); 
		}
		if(!isNull(data.Width)) {
			sheet.setColumnWidth(index, data.Width); 
		}
	}


	private function setRowStyle(required struct data, required any row) localmode=true {
		// row Height
		if(!isNull(row)) {
			if(!isNull(data.HeightInPoints)) {
				data.HeightInPoints = len(data.HeightInPoints)==0 ? 100 : data.HeightInPoints;
				ratio=row.getHeight()/row.getHeightInPoints();
				row.setHeight(data.HeightInPoints*ratio); 
			}
			if(!isNull(data.height)) {
				row.setHeight(data.height); 
			}
		}
	}

	private function getStyle(style) localmode=true {
		if(isNull(style)) return nullValue();
		sct=structNew('linked');
		
		sct['Alignment']=style.getAlignmentEnum().toString();

		sct['FillBackgroundColor']=color=toColor(style.getFillBackgroundColorColor());
		sct['Border']=getBorder(style);
		sct['BorderColor']=getBorderColor(style);

		sct['FillPattern']=style.getFillPatternEnum().toString();
		sct['Font']=getFont(style);
		sct['Font']=getFont(style);
		sct['ForegroundColor']=toColor(style.getFillForegroundColorColor());
		sct['Hidden']=style.getHidden();
		sct['Indention']=style.getIndention();
		sct['Locked']=style.getLocked();
		//sct['QuotePrefixed']=style.getQuotePrefixed();
		sct['ReadingOrder']=getReadingOrder(style);
		sct['Rotation']=style.getRotation();
		sct['ShrinkToFit']=style.getShrinkToFit();
		sct['VerticalAlignmentEnum']=style.getVerticalAlignmentEnum().toString();
		sct['WrapText']=style.getWrapText();

		// HSS
		if("hssf"==getType()) {
			sct['DataFormat']=style.getDataFormatString(poi);
			sct['UserStyleName']=style.getUserStyleName();
		}
		// XSS
		if("xssf"==getType()) {
			sct['DataFormat']=style.getDataFormatString();
		}

		//if(color!='auto')sct['raw']=style;
		return sct;

		// Font.ANSI_CHARSET, Font.DEFAULT_CHARSET, Font.SYMBOL_CHARSET
	}

	private string function getString(required richText) localmode=true {
		//return richText.numFormattingRuns();
		return richText.getString();
	}

	private struct function getBorderColor(required style) localmode=true {
		sct=structNew('linked');
		if(getType()=="xssf") {
			sct['Left']=toColor(style.getLeftBorderXSSFColor());
			sct['Top']=toColor(style.getTopBorderXSSFColor());
			sct['Right']=toColor(style.getRightBorderXSSFColor());
			sct['Bottom']=toColor(style.getBottomBorderXSSFColor());
		}
		else {
			sct['Left']=toColor(style.getLeftBorderColor());
			sct['Top']=toColor(style.getTopBorderColor());
			sct['Right']=toColor(style.getRightBorderColor());
			sct['Bottom']=toColor(style.getBottomBorderColor());
		}

		return sct;
	}
	private void function setBorderColor(required style, required data) localmode=true {
		if(!isStruct(data)) { 
			var str=toString(data);
			data={left:str,right:str,top:str,bottom:str};
		}
		
		if(!isNull(data.left)) style.setLeftBorderColor(toColorShort(data.left));
		if(!isNull(data.top)) style.setTopBorderColor(toColorShort(data.Top));
		if(!isNull(data.right)) style.setRightBorderColor(toColorShort(data.Right));
		if(!isNull(data.bottom)) style.setBottomBorderColor(toColorShort(data.Bottom));
	}
	
	private struct function getBorder(required style) localmode=true {
		sct=structNew('linked');
		sct['Left']=getBorderStyle(style.getBorderLeftEnum());
		sct['Top']=getBorderStyle(style.getBorderTopEnum());
		sct['Right']=getBorderStyle(style.getBorderRightEnum());
		sct['Bottom']=getBorderStyle(style.getBorderBottomEnum());
		return sct;
	}

	private void function setBorder(required style, required data) localmode=true {
		if(!isStruct(data)) {
			var str=toString(data);
			data={left:str,right:str,top:str,bottom:str};
		}
		style.setBorderLeft(toBorderStyle(isNull(data.left)?nullValue():data.left));
		style.setBorderTop(toBorderStyle(isNull(data.Top)?nullValue():data.Top));
		style.setBorderRight(toBorderStyle(isNull(data.Right)?nullValue():data.Right));
		style.setBorderBottom(toBorderStyle(isNull(data.Bottom)?nullValue():data.Bottom));
	}

	private any function getBorderStyle(required borderStyle) localmode=true {
		return borderStyle.toString();
	}

	private any function toBorderStyle(data) localmode=true {
		if(!isNull(data)){
			if("DASH_DOT"==data) return BorderStyle.DASH_DOT;
			else if("DASH_DOT_DOT"==data) return BorderStyle.DASH_DOT_DOT;
			else if("DASHED"==data) return BorderStyle.DASHED;
			else if("DOTTED"==data) return BorderStyle.DOTTED;
			else if("DOUBLE"==data) return BorderStyle.DOUBLE;
			else if("HAIR"==data) return BorderStyle.HAIR;
			else if("MEDIUM"==data) return BorderStyle.MEDIUM;
			else if("MEDIUM_DASH_DOT"==data) return BorderStyle.MEDIUM_DASH_DOT;
			else if("MEDIUM_DASH_DOT_DOT"==data) return BorderStyle.MEDIUM_DASH_DOT_DOT;
			else if("MEDIUM_DASHED"==data) return BorderStyle.MEDIUM_DASHED;
			else if("NONE"==data) return BorderStyle.NONE;
			else if("SLANTED_DASH_DOT"==data) return BorderStyle.SLANTED_DASH_DOT;
			else if("THICK"==data) return BorderStyle.THICK;
			else if("THIN"==data) return BorderStyle.THIN;
			else throw "invalid border style [#data#], valid border styles are [DASH_DOT,DASH_DOT_DOT,DASHED,DOTTED,DOUBLE,HAIR,MEDIUM,MEDIUM_DASH_DOT,MEDIUM_DASH_DOT_DOT,MEDIUM_DASHED,NONE,SLANTED_DASH_DOT,THICK,THIN].";
		}
		return BorderStyle.NONE;
	}

	private void function setFont(required style, required struct data) localmode=true {
		key=createKey(data);
		if(!structKeyExists(variables.fonts,key)) {
			font=variables.poi.createFont();
			
			if(!isNull(data.bold))font.setBold(data.bold==true);
			if(!isNull(data.Italic))font.setItalic(data.Italic==true);
			if(!isNull(data.charset))font.setCharSet(toCharSet(data.charset));
			if(!isNull(data.color))font.setColor(toColorShort(data.color));
			if(!isNull(data.Height))font.setFontHeight(int(data.Height));
			if(!isNull(data.HeightInPoints))font.setFontHeightInPoints(int(data.HeightInPoints));
			if(!isNull(data.name))font.setFontName(data.name);
			if(!isNull(data.Strikeout))font.setStrikeout(data.Strikeout==true);
			if(!isNull(data.TypeOffset))font.setTypeOffset(toTypeOffset(data.TypeOffset));
			if(!isNull(data.Underline))font.setUnderline(toUnderline(data.Underline));
			variables.fonts[key]=font;
		}
		else font=variables.fonts[key];
		style.setFont(font);
	}

	private any function toFillPattern(required string data) localmode=true {
		if(!isNull(data)){
			var patterns="NO_FILL,SOLID_FOREGROUND,FINE_DOTS,ALT_BARS,SPARSE_DOTS,THICK_HORZ_BANDS,THICK_VERT_BANDS,THICK_BACKWARD_DIAG,THICK_FORWARD_DIAG,BIG_SPOTS,BRICKS,THIN_HORZ_BANDS,THIN_VERT_BANDS,THIN_BACKWARD_DIAG,THIN_FORWARD_DIAG,SQUARES,DIAMONDS";

			// we do this because older version of the POI library have the pattern constant in "CellStyle" and newer ave them in "FillPatternType"
			try {
				try {
					return FillPatternType[ucase(data)];
				}
				catch(ee) {
					return CellStyle[ucase(data)];
				}
			}
			catch(e) {
				cfthrow(message="invalid fill pattern [#data#], valid fill patterns are [#patterns#]. ",detail=e.message);
			}

		}
		return style.NO_FILL;
	}

	private struct function getFont(required style) localmode=true {
		if("xssf"==getType()) font=style.getFont();
		else font=style.getFont(poi);
		
		sct['Bold']=font.getBold();
		sct['Italic']=font.getItalic();
		sct['Color']=getFontColor(font);
		sct['Name']=font.getFontName();
		sct['HeightInPoints']=font.getFontHeightInPoints();
		sct['Height']=font.getFontHeight();
		sct['CharSet']=getCharSet(font);
		sct['Strikeout']=font.getStrikeout();
		sct['TypeOffset']=getTypeOffset(font);
		sct['Underline']=getUnderline(font);
		//sct['raw']=font;
		return sct;
	}

	private function getFontColor(required font) localmode=true {
		
		//HSS TODO is this necessary
		try {
			if("hssf"==getType()) 
				return toColor(font.getHSSFColor(poi));
		}
		catch(e) {}
		
		if("xssf"==getType()) return toColor(font.getXSSFColor());
		return toColor(font.getColor());
	}


	// Font.U_NONE, Font.U_SINGLE, Font.U_DOUBLE, Font.U_SINGLE_ACCOUNTING, Font.U_DOUBLE_ACCOUNTING
	private string function getUnderline(required font) localmode=true {
		switch(font.getUnderline()) {
			case font.U_SINGLE: return "SINGLE";
			case font.U_DOUBLE: return "DOUBLE";
			case font.U_SINGLE_ACCOUNTING: return "SINGLE_ACCOUNTING";
			case font.U_DOUBLE_ACCOUNTING: return "DOUBLE_ACCOUNTING";
		}
		return "NONE";
	}

	private string function toUnderline(required string str) localmode=true {
		if("SINGLE"==str) return Font.U_SINGLE;
		if("DOUBLE"==str) return Font.U_DOUBLE;
		if("SINGLE_ACCOUNTING"==str) return Font.U_SINGLE_ACCOUNTING;
		if("DOUBLE_ACCOUNTING"==str) return Font.U_DOUBLE_ACCOUNTING;
		if("NONE"==str) return Font.U_NONE;
		throw "underline value [#str#] is invalid, valid values are [NONE, SINGLE, DOUBLE, SINGLE_ACCOUNTING, DOUBLE_ACCOUNTING]";
	}
	
	// Font., Font.SS_SUPER, Font.SS_SUB
	private string function getTypeOffset(required font) localmode=true {
		switch(font.getTypeOffset()) {
			case font.SS_SUPER: return "SUPER";
			case font.SS_SUB: return "SUB";
		}
		return "NONE";
	}

	private string function toTypeOffset(required string str) localmode=true {
		if("SUPER"==str) return Font.SS_SUPER;
		if("SUB"==str) return Font.SS_SUB;
		if("NONE"==str) return Font.SS_NONE;
		throw "type offset [#str#] is invalid, valid values are [SUPER,SUB,NONE]";
	}

	// Font.ANSI_CHARSET, Font.DEFAULT_CHARSET, Font.SYMBOL_CHARSET
	private string function getCharSet(required font) localmode=true {
		switch(font.getCharSet()) {
			case font.ANSI_CHARSET: return "ANSI_CHARSET";
			case font.SYMBOL_CHARSET: return "SYMBOL_CHARSET";
		}
		return "DEFAULT_CHARSET";
	}

	private string function toCharSet(required string charset) localmode=true {
		if("ANSI_CHARSET"==charset) return Font.ANSI_CHARSET;
		if("SYMBOL_CHARSET"==charset) return Font.SYMBOL_CHARSET;
		throw "invalid charset defintion [#charset#], valid values are [ANSI_CHARSET,SYMBOL_CHARSET]";
	}

	
	private string function toHorizontalAlignment(required str) localmode=true {
		if("CENTER"==str) return HorizontalAlignment.CENTER;
		if("CENTER_SELECTION"==str) return HorizontalAlignment.CENTER_SELECTION;
		if("DISTRIBUTED"==str) return HorizontalAlignment.DISTRIBUTED;
		if("FILL"==str) return HorizontalAlignment.FILL;
		if("GENERAL"==str) return HorizontalAlignment.GENERAL;
		if("JUSTIFY"==str) return HorizontalAlignment.JUSTIFY;
		if("LEFT"==str) return HorizontalAlignment.LEFT;
		if("RIGHT"==str) return HorizontalAlignment.RIGHT;
		throw "Horizontal Alignment value [#str#] is invalid, valid values are [CENTER,CENTER_SELECTION,DISTRIBUTED,FILL,GENERAL,JUSTIFY,LEFT,RIGHT]";
	}	

	private string function toVerticalAlignment(required str) localmode=true {
		if("TOP"==str) return CellStyle.VERTICAL_TOP;
		if("CENTER"==str) return CellStyle.VERTICAL_CENTER;
		if("BOTTOM"==str) return CellStyle.VERTICAL_BOTTOM;
		if("JUSTIFY"==str) return CellStyle.VERTICAL_JUSTIFY;
		throw "Vertical Alignment value [#str#] is invalid, valid values are [TOP,CENTER,BOTTOM,JUSTIFY]";
	}

	private string function getHorizontalAlignment(required comment) localmode=true {
		switch(comment.getHorizontalAlignment()) {
			case comment.HORIZONTAL_ALIGNMENT_LEFT: return "LEFT";
			case comment.HORIZONTAL_ALIGNMENT_CENTERED: return "CENTERED";
			case comment.HORIZONTAL_ALIGNMENT_RIGHT: return "RIGHT";
			case comment.HORIZONTAL_ALIGNMENT_JUSTIFIED: return "JUSTIFIED";
			case comment.HORIZONTAL_ALIGNMENT_DISTRIBUTED: return "DISTRIBUTED";
		}
		return "LEFT";
	}
	private string function getVerticalAlignment(required comment) localmode=true {
		switch(comment.getVerticalAlignment()) {
			case comment.VERTICAL_ALIGNMENT_TOP: return "TOP";
			case comment.VERTICAL_ALIGNMENT_CENTER: return "CENTER";
			case comment.VERTICAL_ALIGNMENT_BOTTOM: return "BOTTOM";
			case comment.VERTICAL_ALIGNMENT_JUSTIFY: return "JUSTIFY";
			case comment.VERTICAL_ALIGNMENT_DISTRIBUTED: return "DISTRIBUTED";
		}
		return "TOP";
	}
	private string function getLineStyle(required comment) localmode=true {
		switch(comment.getLineStyle()) {
			case comment.LINESTYLE_SOLID: return "SOLID";
			case comment.LINESTYLE_DASHSYS: return "DASHSYS";
			case comment.LINESTYLE_DOTSYS: return "DOTSYS";
			case comment.LINESTYLE_DASHDOTSYS: return "DASHDOTSYS";
			case comment.LINESTYLE_DASHDOTDOTSYS: return "";
			case comment.LINESTYLE_DOTGEL: return "DOTGEL";
			case comment.LINESTYLE_DASHGEL: return "DASHGEL";
			case comment.LINESTYLE_LONGDASHGEL: return "LONGDASHGEL";
			case comment.LINESTYLE_DASHDOTGEL: return "DASHDOTGEL";
			case comment.LINESTYLE_LONGDASHDOTGEL: return "LONGDASHDOTGEL";
			case comment.LINESTYLE_LONGDASHDOTDOTGEL: return "LONGDASHDOTDOTGEL";
			case comment.LINESTYLE_NONE: return "NONE";
			case comment.LINESTYLE_DEFAULT: return "DEFAULT";
		}
		return "NONE";
	}

	// 0 means Context (Default), 1 means Left To Right, and 2 means Right to Left
	private string function getReadingOrder(required style) localmode=true {
		if("hssf"==getType()) {
			switch(style.getReadingOrder()) {
				case 1: return "LTR";
				case 2: return "RTL";
			}
			return "DEFAULT";
		}
		// TODO XSSF support for this
		return "TODO";
	}

	private numeric function toReadingOrder(required string str) localmode=true {
		if("DEFAULT"==str) return 0;
		if("LTR"==str) return 1;
		if("RTL"==str) return 2;
		throw "Reading Order [#str#] is invalid, valid values are [LTR, RTL, DEFAULT]";
	}

	private struct function getSpreadsheetVersion() localmode=true {
		sct=structNew("linked");
		sv=poi.getSpreadsheetVersion();
		sct['max']['Columns']=sv.getMaxColumns();
		sct['max']['Rows']=sv.getMaxRows();
		sct['max']['CellStyles']=sv.getMaxCellStyles();
		sct['max']['ConditionalFormats']=sv.getMaxConditionalFormats();
		sct['max']['FunctionArgs']=sv.getMaxFunctionArgs();
		
		sct['last']['ColumnIndex']=sv.getLastColumnIndex();
		sct['last']['ColumnName']=sv.getLastColumnName();
		sct['last']['RowIndex']=sv.getLastRowIndex();
		return sct;
	}
	

	private any function getValue(required cell) localmode=true {
		switch(getCellType(cell)) {
			case "datetime": return cell.getDateCellValue();
			case "numeric": return cell.getNumericCellValue();
			//case cell.CELL_TYPE_STRING: return "string";
			case "formula": return cell.getCellFormula();
			case "blank": return "";
			case "boolean": return cell.getBooleanCellValue();
			//case cell.CELL_TYPE_ERROR: return "error";
		}


		return cell.toString();
	}

	private string function toType(required string dbType, required string defaultValue) localmode=true {
		switch(dbType) {
			case "varchar": return 'string';
			case "date": 
			case "timestamp": 
				return 'datetime';
			case "int": 
			case "integer":
			case "double": 
				return 'number';
		}
		return defaultValue;
	}

	function toCFType(dbType) {
		switch(dbType) {
			case "varchar": return 'string';
			case "date": 
			case "timestamp": 
				return 'datetime';
			case "int": 
			case "integer":
			case "double": 
				return 'number';
		}
	}

	function extractTypes(query qry, boolean fast=true) {
		var orgTypes=[];
		loop array=getMetaData(qry) index="local.i" item="local.el" {
			orgTypes[i]=toCFType(el.typeName);
		}
		if(fast)return orgTypes;

		var columns=queryColumnArray(qry);
		var max=50;
		var types=[];
		
		// define rows we check
		if(qry.recordcount>(max*2)) {
			local.rows=[
				{from:1,to:max}
				,{from:qry.recordcount-max,to:qry.recordcount}
			];
		}
		else {
			local.rows=[
				{from:1,to:qry.recordcount}
			];
		}
		
		// analyze types
		loop array=columns index="local.icol" item="local.col" {
			local.type="datetime";
			loop array=rows item="local._row" {
				loop from=_row.from to=_row.to index="local.row" {
					var val=qry[col][row];

					if(isStruct(val) && !isNull(val.value)) {
						val=val.value;
					}
					if(isSimpleValue(val)) {
						if(val=="" || val=="-" || val=="+") continue;
					}


					// timestamp
					if(type=="datetime") {
						 if(isDate(val)) {
						 	var arr=listToArray(val,':');
						 	if(!(arrayLen(arr)==2 && isNumeric(arr[1]) && isNumeric(arr[2]))) {
								continue;
						 	}
						 }
						 local.type="boolean";
					}
					if(type=="boolean") {
						var tmp=isSimpleValue(val)?("_"&val):"";
						if(tmp=="_true" || tmp=="_false" || tmp=="_yes" || tmp=="_no") continue;
						else local.type="number";

					}
					if(type=="number") {
						if(isNumeric(val)) continue;
						 else local.type="string";
					}
					local.type="string";
				}
			}
			types[icol]=type;
		}
		return types;
	}

	/*private array function getmyOwnMetaData(qry) {
		local.aRet = [];
		local.qQry = arguments.qry;
		if (qQry.recordCount lt 2) return getMetaData(qQry);
		loop query=qQry {
			// toDo
			if (qQry.currentRow eq 1) continue; // the first row will always contain the column headers
			loop list=qQry.columnList index="local.iField" item="local.sField" {
				aRet[iField] = aRet[iField] ?:{"isCaseSensitive":false,"name":sField, "typeName": ""};
				if (isDate(qQry[sField])) {
					local.sType = "date";
					if (qQry[sField] eq int(qQry[sField])) {
						sType = "timestamp";
					}
				} else if (isNumeric(qQry[sField])) {
					sType = "number";
				} else if (isBoolean(qQry[sField])) {
					sType = "numeric";
				} else {
					sType = "string";
				};
				if (len(aRet[iField].typeName)==0) {
					aRet[iField].typeName = sType;
				} else if (aRet[iField].typeName neq sType) {
					aRet[iField].typeName = "string";
				}
			}
			if (qQry.currentRow gt 2) {
				break;
			}
		}
		return aRet;
	}*/

	variables.rootKeys={type:'',value:''};
	private boolean function isExtendedValue(required value) localmode=true {
		if(!isStruct(value) || structCount(value)==0) return false;
		// can have more if(structCount(value)>structCount(variables.rootKeys)) return false;
		loop struct=variables.rootKeys index="k" item="v" {
			if(!structKeyExists(value,k)) {
				return false;
			}
		}
		return true;
	}

	private function getHeap() {
		var q=getMemoryUsage();
		var data.used=0;
		var data.max=0;
		
		loop query=q {
			if(q.type!="HEAP") continue;
			data.used+=q.used;
			data.max+=q.max;
		}

		return 1/data.max*data.used;
	}

	
	
	public static query function addTypes(required query qry, required array types, boolean clone=false) localmode=true {
		if(clone) qry=duplicate(qry);
		columns=qry.columnArray();
		loop from=1 to=qry.recordcount index="row" {
			loop array=columns index="i" item="col" {
				qry[col][row]={type:types[i]?:'string',value:qry[col][row]};
			}
		}
		variables.isNew=false;
		return qry;
	}



	public static function loadPoi(className) {
		return createObject('java',className);
		//return createObject('java',className,"org.lucee.poi.ooxml","3.15.0");
	}
	public static function loadSystem(className) {
		return createObject('java',className);
	}
}