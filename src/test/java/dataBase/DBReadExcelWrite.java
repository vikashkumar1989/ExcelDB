package dataBase;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.AfterClass;
import org.testng.annotations.Test;

public class DBReadExcelWrite {
	
	List<String> name=new ArrayList<String>();
	List<String> id=new ArrayList<String>();
	List<String> dept=new ArrayList<String>();
	List<Integer> age=new ArrayList<Integer>();
	String dir=System.getProperty("user.dir");
	XSSFWorkbook wb;
	XSSFSheet sh;
	XSSFRow rw;
	FileInputStream fis;
	FileOutputStream fos;
	Properties prop;
	
	@Test(priority=0)
	public void MySQLExecute() throws SQLException, IOException {
			fis=new FileInputStream(dir+"\\src\\test\\java\\dataBase\\Connection.properties");
			prop=new Properties();
			prop.load(fis);
		    String url = prop.getProperty("url");
		    String user = prop.getProperty("user");
		    String password = prop.getProperty("password");
		    Connection con = DriverManager.getConnection(url, user, password);
		    Statement smt=con.createStatement();
		    ResultSet selquery=smt.executeQuery("select * from Employeeinfo;");
		    while(selquery.next()){
		    	name.add(selquery.getString("name"));
		    	id.add(selquery.getString("id"));
		    	dept.add(selquery.getString("dept"));
		    	age.add(selquery.getInt("age"));
		    	
		    }
	    
	}
	@Test(priority=1)
	  public void excelWrite() throws IOException { 
		  wb=new XSSFWorkbook(); 
		  sh=wb.createSheet("FromDB");		  
		  rw=sh.createRow(0); 
		  rw.createCell(0).setCellValue("Name");
		  rw.createCell(1).setCellValue("ID");
		  rw.createCell(2).setCellValue("Department");
		  rw.createCell(3).setCellValue("Age");
		  
		  int i=name.size();
		  for(int j=1;j<i;j++) {
			  rw=sh.createRow(j); 
			  rw.createCell(0).setCellValue(name.get(j));
			  rw.createCell(1).setCellValue(id.get(j));
			  rw.createCell(2).setCellValue(dept.get(j));
			  rw.createCell(3).setCellValue(age.get(j));
			   
		  }
		  fos=new FileOutputStream(dir+"\\src\\test\\resources\\Excel\\SQLExecute.xlsx");
		  wb.write(fos);
		  wb.close();
	  
	  }
	@AfterClass
	public void OpenFile() throws IOException {
		Desktop.getDesktop().open(new File(dir+"\\src\\test\\resources\\Excel\\SQLExecute.xlsx"));		
	}
	 

}
