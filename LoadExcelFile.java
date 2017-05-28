import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class LoadConfiguration {
	public String getExcelFile () throws IOException {
		Properties properties = new Properties();
		InputStream inpStream ;
		String fileName = "config.properties";
		
	    try {
	    	inpStream = new FileInputStream(fileName);
			properties.load(inpStream);
		    if (inputStream != null) {
				properties.load(inputStream);
			} else {
				throw new FileNotFoundException("property file '" + fileName + "' not found in the classpath");
			}
		    String excelSheet = prop.getProperty("excel_file");
		} catch (IOException ioException) {
			ioException.printStackTrace();
		}finally {
			inpStream.close();
		}
	    return excelSheet;
	  }
}
