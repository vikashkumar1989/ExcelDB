package dataBase;


import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

public class FileOpen {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String dir=System.getProperty("user.dir");
		Desktop.getDesktop().open(new File(dir+"\\src\\test\\resources\\Excel\\SQLExecute.xlsx"));

	}

}
