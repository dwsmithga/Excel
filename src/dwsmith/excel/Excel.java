package dwsmith.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.ss.util.WorkbookUtil;


//import java.io.File;
import java.io.FileOutputStream;

public class Excel {

	public static void main(String[] args) {
		String file = "Test3.xls";
		String sheetName = "Eggs";
		
		//XSSFWorkbook workbook = new XSSFWorkbook();
		Workbook workbook = new XSSFWorkbook();
		
		//XSSFSheet sheet = workbook.createSheet(sheetName);
		Sheet sheet = workbook.createSheet(sheetName);
		
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(3);
		
		Cell cell1 = sheet.createRow(0).createCell(0);
		Cell cell2 = sheet.createRow(0).createCell(1);		
		Cell cell3 = sheet.createRow(0).createCell(2);
		Cell cell4 = sheet.createRow(0).createCell(3);
		Cell cell5 = sheet.createRow(0).createCell(4);
		
		cell1.setCellValue(46);
		cell2.setCellValue(25);
		cell3.setCellValue(199);
		cell4.setCellValue(34);
		cell5.setCellFormula("sum(a1:d1)");
		
		//cell.setCellValue("Hi there2");
		//System.out.println(cell.getRichStringCellValue().toString());
		
		
		
		try {
			FileOutputStream output = new FileOutputStream(file);
			workbook.write(output);
			output.close();
			workbook.close();
		} catch(Exception e) {
			e.printStackTrace();
		}
		
	}
	
}
