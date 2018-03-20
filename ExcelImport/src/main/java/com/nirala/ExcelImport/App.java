package com.nirala.ExcelImport;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Arrays;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class App 
{
	public static final String SAMPLE_XLSX_FILE_PATH = "E://kunal/user.xlsx";
	static final String JDBC_DRIVER = "com.mysql.jdbc.Driver";  
	static final String DB_URL = "jdbc:mysql://127.0.0.1:3306/kunal";

	   //  Database credentials
	 static final String USER = "kunal";
	 static final String PASS = "kunal@123";
	 
    @SuppressWarnings("deprecation")
	public static void main( String[] args)
    {
    	   Connection conn = null;
    	   Statement stmt = null;
    	   StringBuilder sb = new StringBuilder();
    	   String statement = null;
    	   try{
    	      //STEP 2: Register JDBC driver
    	      Class.forName("com.mysql.jdbc.Driver");

    	      //STEP 3: Open a connection
    	      System.out.println("Connecting to a selected database...");
    	      conn = DriverManager.getConnection(DB_URL, USER, PASS);
    	      System.out.println("Connected database successfully...");
    	      
    	      //STEP 4: Execute a query
    	      System.out.println("Creating table in given database...");
    	      stmt = (Statement) conn.createStatement();
    	     
    	        Workbook workbook = null;
    	        DataFormatter dataFormatter = new DataFormatter();
    			try {
    				workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
    			} catch (EncryptedDocumentException e) {
    				e.printStackTrace();
    			} catch (InvalidFormatException e) {
    				e.printStackTrace();
    			} catch (IOException e) {
    				e.printStackTrace();
    			}
    	        boolean flag = true;
    	        for(Sheet sheet: workbook) {
    	            Iterator<Row> rowIterator = sheet.rowIterator();
    	            int r = 0;
    	            int rid = 1;
    	            int count=1;
    	            while (rowIterator.hasNext()) {
    	                Row row = rowIterator.next();
    	                if(flag){
    	                	Iterator<Cell> cellIterator = row.cellIterator();
	    	                String sql = "CREATE TABLE IF NOT EXISTS "+sheet.getSheetName()+"(id INTEGER not NULL, ";
							 while (cellIterator.hasNext()) {
								 Cell cell = cellIterator.next();
								 sql += dataFormatter.formatCellValue(cell)+" VARCHAR(255)";
								if(cellIterator.hasNext()) sql += ",";
								r++;
							 }
							 sql += ", PRIMARY KEY ( id ) );";
								for(int i=0;i<=r;i++){
									sb.append('?');
									if(i<r)sb.append(',');
								}
							  statement = "insert into "+sheet.getSheetName()+" values("+sb+");";
	    	     	         ((java.sql.Statement) stmt).executeUpdate(sql);
	    	     	         flag = false;
    	                 }else{
        	                PreparedStatement stmt1=conn.prepareStatement(statement); 
							Iterator<Cell> cellIterator = row.cellIterator();
							stmt1.setString(1,""+rid);
							int c = 1;
        	               while (cellIterator.hasNext() && c<9) {
        	            	    c =c+1;
        	                	Cell cell = cellIterator.next();
        	                	if(cell != null){
        	                		String content=dataFormatter.formatCellValue(cell);
        	                		stmt1.setString(c,content);
        	                	}
        	                	else{
        	                		stmt1.setString(c," ");
        	                	}
        	                }
        	                while(c<9){
        	            	   c++;
        	            	   stmt1.setString(c," ");   
        	               }
        	               int tot=stmt1.executeUpdate();  
        	               rid++;
    	                }
    	            }
    	        }  
    	   }catch(SQLException se){
    	      se.printStackTrace();
    	   }catch(Exception e){
    	      e.printStackTrace();
    	   }finally{
    	      try{
    	         if(stmt!=null)
    	            conn.close();
    	      }catch(SQLException se){
    	      }
    	      try{
    	         if(conn!=null)
    	            conn.close();
    	      }catch(SQLException se){
    	         se.printStackTrace();
    	      }
    	   }
     System.out.println("Thanks!");
    }
}
