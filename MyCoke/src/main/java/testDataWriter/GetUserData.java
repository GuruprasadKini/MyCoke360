package testDataWriter;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetUserData {
	 	static Map<String, String> UserData = new HashMap<>();;
	 	static Map<String, String[]> accountData = new HashMap<>();
	 	static String [] username = new String[601];
	    static int RowNum = 1;
	    static FileOutputStream fileOut;
	    static XSSFSheet sheet;
	    
	public static void getCredentials(String filePath, int users) throws IOException {
        try (FileInputStream fileIn = new FileInputStream(filePath)) {
            IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);
            XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
            sheet = workbook.getSheet("UserData");
            DataFormatter formatter = new DataFormatter();
            for (int rowNum = 1; rowNum < users + 1; rowNum++) {
                XSSFRow row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
                XSSFCell cell = row.getCell(0);
                XSSFCell cell1 = row.getCell(1);
                String usernameCellValue = formatter.formatCellValue(cell);
                String currentAccountId = formatter.formatCellValue(cell1);
                username[rowNum - 1] = usernameCellValue;
                UserData.put(usernameCellValue, currentAccountId);
                XSSFSheet currentSheet = workbook.getSheet(currentAccountId);
                XSSFRow currentRow = currentSheet.getRow(0);
                String[] SKUs = new String[50];
	            for (int cellNum = 0; cellNum < 50; cellNum++) {
	               XSSFCell cell3 = currentRow.getCell(cellNum);
	               String currentSKU = formatter.formatCellValue(cell3).toString();
	               SKUs[cellNum] = currentSKU;
	               accountData.put(currentAccountId, SKUs);                	
	            }	 
            }
            System.out.println(accountData);
            workbook.close();
            fileIn.close();
        }
	}
}

