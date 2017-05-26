
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.rbc.tlm.dbconnection.DBConnection;


public class DeleteSheet {
	public static void main(String args[]) throws SQLException{
		DeleteSheet deleteSheet=new DeleteSheet();
		deleteSheet.method();
		System.out.println("aaa");
	}
	
	public void method() throws SQLException{
		String fileName="C:/Shruti/Eclipse/RenameSheets/Outstanding_Items_Count 12-May-2017.xlsx";
		int index = 0;
		int rownum = 0;
 	    int cellnum = 0;
		
	try {
		FileInputStream fileStream = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(fileStream);
		XSSFSheet sheet=workbook.getSheet("TLMResultPreviousday");
		index = workbook.getSheetIndex(sheet);
	    workbook.removeSheetAt(index);
	    workbook.setSheetName(workbook.getSheetIndex("TLMResultCurrentday"),("TLMResultPreviousday"));
	    workbook.createSheet("TLMResultCurrentday");
		
	    
	    DBConnection dbConnection=new DBConnection();
	    dbConnection.getConnectionTLM();
	    PreparedStatement preparedStatement= dbConnection.getConnectionTLM().prepareStatement("select count(item_id),b.local_acc_no from ubp0_tlm_dbo.item i,ubp0_tlm_dbo.bank b where i.corr_acc_no = b.corr_acc_no and flag_2 = 0"); 
	    ResultSet result = preparedStatement.executeQuery();
	    
	    while(result.next()){
	    	XSSFSheet sheet1=workbook.getSheet("TLMResultCurrentday");
	    	 
	 	   Row row = sheet1.createRow(rownum++);
	 	   for(int j=0;j<1;j++){
	        	Cell cell = row.createCell(cellnum++);
	            cell.setCellValue(result.getString(j));
	        } 
	    }
	    	    
	    
	    //writing to excel
	    
	  
		FileOutputStream output = new FileOutputStream(fileName);
		workbook.write(output);
		output.close();
		
		
		workbook.close();   
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	
	
	
}
}
