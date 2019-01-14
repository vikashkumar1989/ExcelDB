package dataBase;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.AfterClass;
import org.testng.annotations.Test;

public class ExcelReadDBWrite {
	List<String> name=new ArrayList<String>();
	List<String> id=new ArrayList<String>();
	List<String> dept=new ArrayList<String>();
	List<Integer> age=new ArrayList<Integer>();
	String dir=System.getProperty("user.dir");
	XSSFWorkbook wb;
	XSSFSheet sh;
	XSSFRow rw;
	FileInputStream fis;
	Properties prop;
	
	@Test(priority=1)
	public void MySQLInsert() throws IOException, SQLException {
		fis=new FileInputStream(dir+"\\src\\test\\java\\dataBase\\Connection.properties");
		prop=new Properties();
		prop.load(fis);
	    String url = prop.getProperty("url");
	    String user = prop.getProperty("user");
	    String password = prop.getProperty("password");
	    Connection con = DriverManager.getConnection(url, user, password);
	    Statement smt=con.createStatement();
	    int i=name.size();
	    for(int j=0;j<i;j++) {
	    	String name1=name.get(j);
	    	String id1=id.get(j);
	    	String dept1=dept.get(j);
	    	Integer age1=age.get(j);
	    	smt.executeUpdate("INSERT INTO Employeeinfo VALUES(\""+name1+"\",\""+id1+"\",\""+dept1+"\","+age1+");");
	    }
	    
	    
	    	
	}
	@Test(priority=0)
	public void ExcelRead() throws IOException {
		fis=new FileInputStream(dir+"\\src\\test\\resources\\Excel\\SQLUpdate.xlsx");
		wb=new XSSFWorkbook(fis);
		sh=wb.getSheet("ToDB");
		int i=sh.getLastRowNum()-sh.getFirstRowNum();
		for(int j=1;j<=i;j++) {			
			name.add(sh.getRow(j).getCell(0).getStringCellValue());
			id.  add(sh.getRow(j).getCell(1).getStringCellValue());
			dept.add(sh.getRow(j).getCell(2).getStringCellValue());
			age. add((int) sh.getRow(j).getCell(3).getNumericCellValue());			
		}
		System.out.println(name);
		System.out.println(id);
		System.out.println(dept);
		System.out.println(age);
		
	
	}
	@AfterClass
	public void OpenFile() throws IOException {
		Desktop.getDesktop().open(new File(dir+"\\src\\test\\resources\\Excel\\SQLUpdate.xlsx"));		
	}
	

}
