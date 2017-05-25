import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class TestMain {


	
	public static void main(String[] args) {
		Workbook book = null;
		try
		{
			book = new HSSFWorkbook(new FileInputStream("C://Users//akanchha.jaiswal//Downloads//TestExcel.xls"));
		}
		catch (FileNotFoundException e1)
		{
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		catch (IOException e1)
		{
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		Sheet sheet = book.getSheet("Birthdays2");
		Sheet sheet2 = book.getSheet("AKANCHHA");
		Row r = sheet2.getRow(0);
		System.out.println("Row value : "+r.getCell(0).getStringCellValue());
		sheet2.removeRow(r);
		System.out.println("Row removed");
		// first row start with zero 
		Row row = sheet.createRow(0); 
		// we will write name and birthdates in two columns
		// name will be String and birthday would be Date 
		// formatted as dd.mm.yyyy 
		Cell name = row.createCell(0); 
		name.setCellValue("John");
		Cell birthdate = row.createCell(1); 
		// steps to format a cell to display date value in Excel 
		// 1. Create a DataFormat // 2. Create a CellStyle 
		// 3. Set format into CellStyle 
		// 4. Set CellStyle into Cell 
		// 5. Write java.util.Date into Cell 
		DataFormat format = book.createDataFormat();
		CellStyle dateStyle = book.createCellStyle(); 
		dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
		birthdate.setCellStyle(dateStyle); 
		// It's very trick method, deprecated, don't use // year is from 1900, month starts with zero 
		birthdate.setCellValue(new Date(110, 10, 10)); 
		// auto-resizing columns sheet.autoSizeColumn(1); 
		// Now, its time to write content of Excel into File
		try
		{
			book.write(new FileOutputStream("C://Users//akanchha.jaiswal//Downloads//TestExcel.xls"));
			System.out.println("Written in file");
			book.close();
		}
		catch (FileNotFoundException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		catch (IOException e)
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	


	
	
}

}
