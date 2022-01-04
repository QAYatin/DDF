package utility;

import java.io.FileInputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Properties;

public class GeneralUtils 
{
	
	Properties po;
	public Properties readPropertiesFile(String path) throws Exception
	{
		FileInputStream fi = new FileInputStream(path);
		po = new Properties();
		po.load(fi);
		return po;
	}
	
	public String customTimeStamp(String dateTimeFormat)
	{
		LocalDateTime id = LocalDateTime.now();
		DateTimeFormatter df = DateTimeFormatter.ofPattern(dateTimeFormat);
		String formattedDT = id.format(df);
		return formattedDT;
	}
}
