package com.vk.excel;

import java.io.File;  
import java.io.FileInputStream;  
import java.util.Iterator;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
public class ReaadExcelDemo3  
{  
public static void main(String[] args)   
{  
try  
{  
File file = new File("C:\\Users\\VK\\Desktop\\ApiDesignPattern.xlsx");   //creating a new file instance  
FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
//creating Workbook instance that refers to .xlsx file  
XSSFWorkbook wb = new XSSFWorkbook(fis);   
XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
while (itr.hasNext())                 
{  
Row row = itr.next();  
Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
while (cellIterator.hasNext())   
{  
Cell cell = cellIterator.next();  

if(cell.getRowIndex() == 0) {
switch (cell.getCellType())               
{  
case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
System.out.print("<th>"+cell.getStringCellValue().trim()+"</th>");  
break;  
case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
System.out.print("<th>"+cell.getNumericCellValue()+"</th>");  
break;  
default:  
} 
}
else {
	switch (cell.getCellType())               
	{  
	case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
	System.out.print("<td>"+cell.getStringCellValue().trim()+"</td>");  
	break;  
	case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
	System.out.print("<td>"+cell.getNumericCellValue()+"</td>");  
	break;  
	default:  
	} 
}
}  
System.out.println("------------------------------");  
}  
}  
catch(Exception e)  
{  
e.printStackTrace();  
}  
}  
}  