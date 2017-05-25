import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


public class TestMain {


	
	 public static void main(String[] args) throws IOException {
		 TestMain excel = new TestMain();
	        excel.process("C://Users//akanchha.jaiswal//Downloads//Outstanding_Items_Count 12-May-2017.xls");
	    }
	 
	 public void process(String fileName) throws IOException {
	        BufferedInputStream bis = new BufferedInputStream(new FileInputStream(fileName));
	        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
	        HSSFWorkbook myWorkBook = new HSSFWorkbook();
	        HSSFSheet sheet = null;
	        HSSFSheet sheet2 = null;
	        HSSFRow row = null;
	        HSSFCell cell = null;
	        HSSFSheet mySheet = null;
	        HSSFRow myRow = null;
	        HSSFCell myCell = null;
	        int sheets = workbook.getNumberOfSheets();
	        int fCell = 0;
	        int lCell = 0;
	        int fRow = 0;
	        int lRow = 0;
	        for (int iSheet = 2; iSheet < sheets; iSheet++) {
	            sheet = workbook.getSheetAt(iSheet);
	            sheet2 = workbook.getSheetAt(iSheet-1);
	            String sheetName = sheet2.getSheetName();
	            if (sheet != null) {
	                workbook.removeSheetAt(iSheet-1);
	                mySheet = workbook.createSheet(sheetName);

	                fRow = sheet.getFirstRowNum();
	                lRow = sheet.getLastRowNum();
	                for (int iRow = fRow; iRow <= lRow; iRow++) {
	                    row = sheet.getRow(iRow);
	                    myRow = mySheet.createRow(iRow);
	                    if (row != null) {
	                        fCell = row.getFirstCellNum();
	                        lCell = row.getLastCellNum();
	                        for (int iCell = fCell; iCell < lCell; iCell++) {
	                            cell = row.getCell(iCell);
	                            myCell = myRow.createCell(iCell);
	                            if (cell != null) {
	                                myCell.setCellType(cell.getCellType());
	                                switch (cell.getCellType()) {
	                                //case HSSFCell.CELL_TYPE_BLANK:
	                                    //myCell.setCellValue("");
	                                   // break;

	                                //case HSSFCell.CELL_TYPE_BOOLEAN:
	                                    //myCell.setCellValue(cell.getBooleanCellValue());
	                                   // break;

	                               // case HSSFCell.CELL_TYPE_ERROR:
	                                   // myCell.setCellErrorValue(cell.getErrorCellValue());
	                                  //  break;

	                                //case HSSFCell.CELL_TYPE_FORMULA:
	                                    //myCell.setCellFormula(cell.getCellFormula());
	                                   // break;

	                                case HSSFCell.CELL_TYPE_NUMERIC:
	                                    myCell.setCellValue(cell.getNumericCellValue());
	                                    break;

	                                case HSSFCell.CELL_TYPE_STRING:
	                                    myCell.setCellValue(cell.getStringCellValue());
	                                    break;
	                                default:
	                                    myCell.setCellFormula(cell.getCellFormula());
	                                }
	                            }
	                        }
	                    }
	                }
	            }
	             
	           
	        }
	        
	        myWorkBook.createSheet("TLMResultCurrentday");
	        //Now update the current sheet with the data base values
	        bis.close();
	        BufferedOutputStream bos = new BufferedOutputStream(
	                new FileOutputStream("C://Users//akanchha.jaiswal//Downloads//Outstanding_Items_Count 12-May-2017.xls", false));
	        workbook.write(bos);
	        bos.close();
	    }
	}

