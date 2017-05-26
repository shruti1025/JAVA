import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;


public class TestMain {


	
	 public static void main(String[] args) throws IOException {
		 TestMain excel = new TestMain();
	        excel.process("C://Users//akanchha.jaiswal//Downloads//Outstanding_Items_Count 12-May-2017.xls");
	    }
	 
	 public void process(String fileName) throws IOException {
	        BufferedInputStream bis = new BufferedInputStream(new FileInputStream(fileName));
	        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
	        HSSFWorkbook myWorkBook = new HSSFWorkbook();

	            
	           
	       workbook.removeSheetAt(workbook.getSheetIndex("TLMResultPreviousday"));
	       workbook.setSheetName(workbook.getSheetIndex("TLMResultCurrentday"), "TLMResultPreviousday");
	       workbook.createSheet("TLMResultCurrentday");
	        //Now update the current sheet with the data base values
	       ArrayList<String[]> dataList = new ArrayList<String[]>();
	   

	       
	       Connection conn = null;
	       Statement stmt = null;
	       try{
	        
	          Class.forName("com.mysql.jdbc.Driver");

	          conn = DriverManager.getConnection("sdfsdf","sdfsdf","adasd");

	         
	          stmt = conn.createStatement();
	          String sql;
	          sql = "SELECT id,name FROM Employees";
	          ResultSet rs = stmt.executeQuery(sql);

	        
	          while(rs.next()){
	        //	  rs.getString(0), rs.getString(1)
	        	  String[] s = new String[2];
	        	  s[0] = rs.getString(0);
	        	  s[1] = rs.getString(1);
	        	  dataList.add(s);
	          }
	        
	          rs.close();
	          stmt.close();
	          conn.close();
	       }catch(SQLException se){
	          
	          se.printStackTrace();
	       }catch(Exception e){
	          //Handle errors for Class.forName
	          e.printStackTrace();
	       }finally{
	          //finally block used to close resources
	          try{
	             if(stmt!=null)
	                stmt.close();
	          }catch(SQLException se2){
	          }// nothing we can do
	          try{
	             if(conn!=null)
	                conn.close();
	          }catch(SQLException se){
	             se.printStackTrace();
	          }//end finally try
	       }
	        for(int i=0;i<dataList.size();i++){
	        	Sheet sheet = workbook.getSheet("TLMResultCurrentday");
	        	Row row = sheet.createRow(i);
	        	String[] s = dataList.get(i);
	        	Cell cell1 = row.createCell(0);
	        	Cell cell2 = row.createCell(1);
	        	cell1.setCellValue(s[0]);
	        	cell2.setCellValue(s[1]);
	        }
	        bis.close();
	        BufferedOutputStream bos = new BufferedOutputStream(
	                new FileOutputStream("C://Users//akanchha.jaiswal//Downloads//Outstanding_Items_Count 12-May-2017.xls", false));
	        workbook.write(bos);
	        bos.close();
	        }
	    }
	 

	