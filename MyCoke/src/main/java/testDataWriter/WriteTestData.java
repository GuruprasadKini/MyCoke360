package testDataWriter;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.Random;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteTestData {
    private static final String TEST_DATA_FILE_PATH = "C:/Users/PC-1/eclipse-workspace/MyCoke/src/main/resources/TestData.xlsx";
    private static final String USER_DATA_FILE_PATH = "C:/Users/PC-1/eclipse-workspace/MyCoke/src/main/resources/UserData.xlsx";
    static int RowNum = 1;
    static FileOutputStream fileOut;
    static XSSFWorkbook workbook2;
    static XSSFSheet sheet;
    static String[] SKUs;

	 public static void writeUserData(int totalRows, String access_Token) throws IOException {
	        try (FileOutputStream fileOut = new FileOutputStream(TEST_DATA_FILE_PATH)) {
	            workbook2 = new XSSFWorkbook();
	            XSSFSheet sheet2 = workbook2.createSheet();
	            
	    		//Set markers for Line items 
	    		//50% of items are between 1-10
	    		int lowerWeightage = (int) (totalRows*0.5);
		    		int two = lowerWeightage/4; 
		    		int five = two + lowerWeightage/4;
		    		int eight = five + lowerWeightage/4;
		    		int ten = eight + lowerWeightage/4;
	    		//30% of items are between 11-20
		    	int midWeightage = (int) (totalRows*0.3);
		    		int twelve = ten + midWeightage/3;
		    		int fifteen = twelve + midWeightage/3;
		    		int eighteen = fifteen + midWeightage/3;
	    		//10% of items are 20+
		    	int topWeightage = (int) (totalRows*0.2);
		    		int twentyfive = eighteen + topWeightage/2;
		    		int thirty = twentyfive + topWeightage/2;

	            // Create the header row outside the loop
	            XSSFRow headerRow = sheet2.createRow(0);
	            String[] headerNames = {"Email","AdminToken", "products","Product1","Product2","Product3","Product4","Product5","Product6","Product7","Product8","Product9","Product10","Product11","Product12","Product13","Product14","Product15","Product16","Product17","Product18","Product19","Product20","Product21","Product22","Product23","Product24","Product25","Product26","Product27","Product28","Product29","Product30","Quantity1","Quantity2","Quantity3","Quantity4","Quantity5","Quantity6","Quantity7","Quantity8","Quantity9","Quantity10","Quantity11","Quantity12","Quantity13","Quantity14","Quantity15","Quantity16","Quantity17","Quantity18","Quantity19","Quantity20","Quantity21","Quantity22","Quantity23","Quantity24","Quantity25","Quantity26","Quantity27","Quantity28","Quantity29","Quantity30","UpdateQty1","UpdateQty2","UpdateQty3","UpdateQty4","UpdateQty5","UpdateQty6","UpdateQty7","UpdateQty8","UpdateQty9","UpdateQty10","UpdateQty11","UpdateQty12","UpdateQty13","UpdateQty14","UpdateQty15","UpdateQty16","UpdateQty17","UpdateQty18","UpdateQty19","UpdateQty20","UpdateQty21","UpdateQty22","UpdateQty23","UpdateQty24","UpdateQty25","UpdateQty26","UpdateQty27","UpdateQty28","UpdateQty29","UpdateQty30"};
                int index = 0;
	            for (int cellNum = 0; cellNum < 93; cellNum++) {
	                XSSFCell headerCell = headerRow.createCell(cellNum);
	                headerCell.setCellValue(headerNames[cellNum]);
	            }

	            // Loop through rows
	            for (int rownumber = 1; rownumber <= totalRows ; rownumber++) {
	                XSSFRow row = sheet2.createRow(rownumber);
	                Random random = new Random();
	                String currentUsername = GetUserData.username[rownumber - 1];
//	                System.out.println(currentUsername);
	                String currentAccountId = GetUserData.UserData.get(currentUsername);
	                SKUs = GetUserData.accountData.get(currentAccountId);
//	            	System.out.println(SKUs[0]);
	                index++;
	                // Loop through cells
	                for (int cellNum = 0; cellNum < 93; cellNum++) {
	                    XSSFCell cell = row.createCell(cellNum);

	                    if (cellNum == 0) {
	                        cell.setCellValue(currentUsername);
	                    }
	                    else if(cellNum == 1){
	                    	cell.setCellValue(access_Token);
	                    }
	                    else if (cellNum == 2) {
	                    	if(rownumber <= two) {
	                    		cell.setCellValue("2");
	                    	}
	                    	if(rownumber > two && rownumber <= five) {
	                    		cell.setCellValue("5");
	                    	}
	                    	if(rownumber > five && rownumber <= eight) {
	                    		cell.setCellValue("8");
	                    	}
	                    	if(rownumber > eight && rownumber <= ten) {
	                    		cell.setCellValue("10");
	                    	}
	                    	if(rownumber > ten && rownumber <= twelve) {
	                    		cell.setCellValue("12");
	                    	}
	                    	if(rownumber > twelve && rownumber <= fifteen) {
	                    		cell.setCellValue("15");
	                    	}
	                    	if(rownumber > fifteen && rownumber <= eighteen) {
	                    		cell.setCellValue("18");
	                    	}
	                    	if(rownumber > eighteen && rownumber <= twentyfive) {
	                    		cell.setCellValue("25");
	                    	}
	                    	if(rownumber > twentyfive && rownumber <= thirty) {
	                    		cell.setCellValue("30");
	                    	}
	                    }
	                    else if (cellNum > 2 && cellNum < 33) {
	                    		cell.setCellValue(SKUs[cellNum - 2]);
	                    }
	                    else if (cellNum > 32 && cellNum < 93) {
	                        cell.setCellValue(random.nextInt(10) + 1);
	                    }
	                }
	            }
	            workbook2.write(fileOut);
	            fileOut.close();
	            workbook2.close();
	        }
	    }
	
	 public static void excelToCsv() throws IOException {
  	   FileInputStream fileIn = new FileInputStream(TEST_DATA_FILE_PATH);
         @SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
         XSSFSheet sheet = workbook.getSheetAt(0);
         fileIn.close();

         // Write the CSV file
         BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(
             new FileOutputStream("C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.csv"), StandardCharsets.UTF_8));
         writer.write('\ufeff'); // add BOM for Excel compatibility

         for (Row row : sheet) {
             for (int i = 0; i < row.getLastCellNum(); i++) {
                 Cell cell = row.getCell(i);
                 if (cell == null) {
                     writer.write("");
                 } else if (cell.getCellType() == CellType.NUMERIC) {
                     writer.write(String.valueOf(cell.getNumericCellValue()));
                 } else if (cell.getCellType() == CellType.STRING) {
                     writer.write(cell.getStringCellValue());
                 }
                 writer.write(",");
             }
             writer.newLine();
         }
         
         writer.flush();
         writer.close();
     }
	 
	 public static void main(String[] args) {
		 try {
			 Scanner threads = new Scanner(System.in);
	    	 System.out.print("Enter number of virtual users for MyCoke360 Performance Test for this system: ");
	    	 int totalRows = (int)(threads.nextInt());; // Set the total number of rows
			 GetUserData.getCredentials(USER_DATA_FILE_PATH, totalRows);
			 GetUserData.getAccessToken();
			 writeUserData(totalRows, GetUserData.access_Token);
			 excelToCsv();	
			 System.out.print("Your data is in the file path - C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.csv");
		 }
		 catch(Exception e){
			 e.printStackTrace();
		 }
	 }
}
