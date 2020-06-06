package com.vk.excel;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelDemo4 {
	
	public static void main(String[] args) {
		try {
			File file = new File("C:\\Users\\VK\\Desktop\\ApiDesignPattern.xlsx");   //creating a new file instance 
			FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
			//creating Workbook instance that refers to .xlsx file  
			XSSFWorkbook wb = new XSSFWorkbook(fis);   
			//System.out.println(wb);
			StringBuilder contentBuilder = new StringBuilder();
			
			 contentBuilder.append("<table>");
			//System.out.print("<table>");
			XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
			for (Row row : sheet) {
				//System.out.println("row no."+row.getRowNum());
				//System.out.println(row.getRowNum());
				
				contentBuilder.append("<tr>");
				//System.out.print("<tr>");
			   for(Cell cell : row)	
			   {
				  
				  // System.out.println("row index"+cell.getRowIndex());
				 //  System.out.println("column index"+cell.getColumnIndex());
				  // System.out.println(cell.getStringCellValue());
				 //  System.out.println(cell.getStringCellValue());
					if (cell.getRowIndex() == 0) {
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING: // field that represents string cell type
							contentBuilder.append("<th>"+cell.getStringCellValue().trim() + "</th>");
							//System.out.print("<th>"+cell.getStringCellValue().trim() + "</th>");
							break;
						case Cell.CELL_TYPE_NUMERIC: // field that represents number cell type
							contentBuilder.append("<th>"+String.valueOf(cell.getNumericCellValue()).trim() + "</th>");
							//System.out.print("<th>"+String.valueOf(cell.getNumericCellValue()).trim() + "</th>");
							break;
						default:
						}

					}
					else
					{
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_STRING: // field that represents string cell type
							
							if((cell.getStringCellValue().trim()).equals("") || (cell.getStringCellValue().trim()).equals(null))
							{
								
							}
							else {
								contentBuilder.append("<td>"+cell.getStringCellValue().trim() + "</td>");
								//System.out.print("<td>"+cell.getStringCellValue().trim() + "</td>");
							}
							
							break;
						case Cell.CELL_TYPE_NUMERIC: // field that represents number cell type
							contentBuilder.append("<td>"+String.valueOf(cell.getNumericCellValue()).trim()+ "</td>");
							//System.out.print("<td>"+String.valueOf(cell.getNumericCellValue()).trim()+ "</td>");
							break;
						default:
						}
					}
			   }
			   contentBuilder.append("</tr>");
			   //System.out.println("</tr>");
			}
			contentBuilder.append("</table>");
			//System.out.println("</table>");
			String content = contentBuilder.toString();
			System.out.println(content);
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	
	
}
