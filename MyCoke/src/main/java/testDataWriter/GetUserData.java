package testDataWriter;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetUserData {
	 	static String[] username = new String[1800];
	    static String[] SKU = new String[30];
	    static int RowNum = 1;
	    static FileOutputStream fileOut;
	    static XSSFSheet sheet;
	    
	public static void getCredentials(String filePath) throws IOException {
        try (FileInputStream fileIn = new FileInputStream(filePath)) {
            IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
            XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
            sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum() + 1; rowNum++) {
                XSSFRow row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
                XSSFCell cell = row.getCell(0);
                username[rowNum - 1] = formatter.formatCellValue(cell).toString();
            }
            sheet = workbook.getSheetAt(1);
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum() + 1; rowNum++) {
                XSSFRow row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
                XSSFCell cell = row.getCell(0);
                SKU[rowNum - 1] = formatter.formatCellValue(cell).toString();
            } 
            workbook.close();
            fileIn.close();
        }
    }
}
