package org.epragati.excel.service;

import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public interface ExcelService {

	/**
	 * Prepare List of Headers
	 * @param headers
	 */
	public void setHeaders(List<String> headers, String key);
	
	/**
	 * Render Result into Excel 
	 * @param result
	 * @return 
	 */
	public XSSFWorkbook renderData(List<List<CellProps>> result, List<String> headers,String fileName,String sheetName);
	
}
