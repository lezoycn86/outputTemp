package com.cnbi.lzytemp.utils;

import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings;
import org.docx4j.openpackaging.parts.SpreadsheetML.TablePart;
import org.docx4j.utils.FoNumberFormatUtil;
import org.xlsx4j.jaxb.Context;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.STCellType;

import java.io.File;
import java.io.InputStream;


public class XSLXUtils {
	
	/** 
	 * 功能描述：创建文档处理包对象 
	 * @return  返回值：返回文档处理包对象 
	 * @throws Exception 
	 * @author cnbi 
	 */
	public static SpreadsheetMLPackage createExcelMLPackage() throws Exception {
		return SpreadsheetMLPackage.createPackage();
	}
	
	/**
	 * 功能描述：加载文档信息 
	 * @param xslxFile
	 * @param password
	 * @return
	 * @throws Exception
	 */
	public static SpreadsheetMLPackage loadExcelPackage(File xslxFile, String password) throws Exception {
		return (SpreadsheetMLPackage) BaseUtils.loadOpcPackage(xslxFile, password);
	}
	
	/**
	 * 功能描述：加载文档信息 
	 * @param input
	 * @param password
	 * @return
	 * @throws Exception
	 */
	public static SpreadsheetMLPackage loadExcelPackage(InputStream input, String password) throws Exception {
		return (SpreadsheetMLPackage) BaseUtils.loadOpcPackage(input, password);
	}
	
	
	/** 
	 * 功能描述：保存文档信息 
	 * @param excelPackage  文档处理包对象
	 * @param fileName     完整的输出文件名称，包括路径 
	 * @throws Exception 
	 * @author cnbi 
	 */
	public static void saveExcelPackage(SpreadsheetMLPackage excelPackage, String fileName) throws Exception {
		saveExcelPackage(excelPackage, new File(fileName));
	}
	
	/** 
	 * 功能描述：保存文档信息 
	 * @param excelPackage  文档处理包对象
	 * @param file         文件 
	 * @throws Exception 
	 * @author cnbi 
	 */
	public static void saveExcelPackage(SpreadsheetMLPackage excelPackage, File file) throws Exception {
		BaseUtils.savePackage(excelPackage, file);
	}
	
	/**
	 * 
	 * @param excelPackage
	 * @return
	 * @throws InvalidFormatException 
	 */
	public static SharedStrings getSharedStrings(SpreadsheetMLPackage excelPackage) throws InvalidFormatException{
		return (SharedStrings) excelPackage.getParts().get(new PartName("/xl/sharedStrings.xml"));
	}
	
	/**
	 * 
	 * @param excelPackage
	 * @param table
	 * @return
	 * @throws InvalidFormatException
	 */
	public static TablePart getTable(SpreadsheetMLPackage excelPackage, String table) throws InvalidFormatException{
		return (TablePart) excelPackage.getParts().get(new PartName("/xl/tables/"+table+".xml"));
	}
	
	
	public static Cell createCell(int rowindex, int colindex, String content ) {
		if(content == null){
			return null;
		}
		Cell cell = Context.getsmlObjectFactory().createCell();
		
		if(Docx4jToolUtils.isNumb(content)){
			cell.setT(STCellType.N);
		}else{
			cell.setT(STCellType.STR);
		}
      String col = FoNumberFormatUtil.format(colindex + 1, FoNumberFormatUtil.FO_PAGENUMBER_UPPERALPHA);
      cell.setR( col+(rowindex+1)); 
		cell.setV(content);
		return cell;
	}
}
