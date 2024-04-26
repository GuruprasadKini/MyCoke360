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
    private static final String PRELIMINARY_DATA_FILE_PATH = "C:/Users/PC-1/eclipse-workspace/MyCoke/src/main/resources/PreliminaryData.xlsx";
    static int RowNum = 1;
    static FileOutputStream fileOut;
    static XSSFWorkbook workbook2;
    static XSSFSheet sheet;

	 public static void writeUserData() throws IOException {
	        try (FileOutputStream fileOut = new FileOutputStream(TEST_DATA_FILE_PATH)) {
	            workbook2 = new XSSFWorkbook();
	            XSSFSheet sheet2 = workbook2.createSheet();
	            Scanner threads = new Scanner(System.in);
	    		System.out.print("Enter number of virtual users for MyCoke360 Performance Test for this system: ");
	    		int totalRows = (int)(threads.nextInt()*24);; // Set the total number of rows
	    		//Set markers for Line items 
	    		//20% of items are 20+
	    		int lowerWeightage = (int) (totalRows*0.2);
		    		int thirty = lowerWeightage/4;
		    		int twentyfive = thirty + lowerWeightage/4;
	    		//30% of items are between 11-20
		    	int midWeightage = lowerWeightage + (int) (totalRows*0.3);
		    		int eighteen = midWeightage/4;
		    		int fifteen = eighteen + midWeightage/4;
		    		int twelve = fifteen + midWeightage/4;
	    		//50% of items are between 1 -10
		    	int topWeightage = midWeightage + (int) (totalRows*0.5);
		    		int ten = topWeightage/4;
		    		int eight = ten + topWeightage/4;
		    		int five = eight + topWeightage/4;
		    		int two = five + topWeightage/4;

	            // Create the header row outside the loop
	            XSSFRow headerRow = sheet2.createRow(0);
	            String[] headerNames = {"Email","products","Product1","Product2","Product3","Product4","Product5","Product6","Product7","Product8","Product9","Product10","Product11","Product12","Product13","Product14","Product15","Product16","Product17","Product18","Product19","Product20","Product21","Product22","Product23","Product24","Product25","Product26","Product27","Product28","Product29","Product30","Quantity1","Quantity2","Quantity3","Quantity4","Quantity5","Quantity6","Quantity7","Quantity8","Quantity9","Quantity10","Quantity11","Quantity12","Quantity13","Quantity14","Quantity15","Quantity16","Quantity17","Quantity18","Quantity19","Quantity20","Quantity21","Quantity22","Quantity23","Quantity24","Quantity25","Quantity26","Quantity27","Quantity28","Quantity29","Quantity30","UpdateQty1","UpdateQty2","UpdateQty3","UpdateQty4","UpdateQty5","UpdateQty6","UpdateQty7","UpdateQty8","UpdateQty9","UpdateQty10","UpdateQty11","UpdateQty12","UpdateQty13","UpdateQty14","UpdateQty15","UpdateQty16","UpdateQty17","UpdateQty18","UpdateQty19","UpdateQty20","UpdateQty21","UpdateQty22","UpdateQty23","UpdateQty24","UpdateQty25","UpdateQty26","UpdateQty27","UpdateQty28","UpdateQty29","UpdateQty30"};
                int index = 0;
	            for (int cellNum = 0; cellNum < 92; cellNum++) {
	                XSSFCell headerCell = headerRow.createCell(cellNum);
	                headerCell.setCellValue(headerNames[cellNum]);
	            }

	            // Loop through rows
	            for (int rownumber = 1; rownumber <= totalRows; rownumber++) {
	                XSSFRow row = sheet2.createRow(rownumber);
	                Random random = new Random();
	                
	                if(index > 600) {
	                	index = 0;
	                }
	                
	                String currentUsername = GetUserData.username[(index)];
	                
	                // Loop through cells
	                for (int cellNum = 0; cellNum < 92; cellNum++) {
	                    XSSFCell cell = row.createCell(cellNum);

	                    if (cellNum == 0) {
	                        cell.setCellValue(currentUsername);
	                        
	                    }
	                    else if (cellNum == 1) {
	                    	if(rownumber <= twentyfive) {
	                    		cell.setCellValue("2");
	                    	}
	                    	if(rownumber <= thirty) {
	                    		cell.setCellValue("30");
	                    	}
	                    	if(rownumber > thirty && rownumber <= twentyfive) {
	                    		cell.setCellValue("25");
	                    	}
	                    	if(rownumber > twentyfive && rownumber <= eighteen) {
	                    		cell.setCellValue("18");
	                    	}
	                    	if(rownumber > eighteen && rownumber <= fifteen) {
	                    		cell.setCellValue("15");
	                    	}
	                    	if(rownumber > fifteen && rownumber <= twelve) {
	                    		cell.setCellValue("12");
	                    	}
	                    	if(rownumber > twelve && rownumber <= ten) {
	                    		cell.setCellValue("10");
	                    	}
	                    	if(rownumber > ten && rownumber <= eight) {
	                    		cell.setCellValue("8");
	                    	}
	                    	if(rownumber > eight && rownumber <= five) {
	                    		cell.setCellValue("5");
	                    	}
	                    	if(rownumber > five && rownumber <= two) {
	                    		cell.setCellValue("2");
	                    	}
	                    }
	                    else if (cellNum > 1 && cellNum < 32) {
	                    	cell.setCellValue(GetUserData.SKU[cellNum - 2]);
	                    }
	                    else if (cellNum > 31 && cellNum < 92) {
	                        cell.setCellValue(random.nextInt(10) + 1);
	                    }
	                }
	            }
	            workbook2.write(fileOut);
	            fileOut.close();
	            workbook2.close();
	        }
	    }
	
//	 public static void ExcelDataCopier() {
//		// Input and output file paths
//	        String inputFile = PRELIMINARY_DATA_FILE_PATH;
//	        String outputFile = TEST_DATA_FILE_PATH;
//	        
//	        try {
//	            // Load the input Excel file
//	            FileInputStream fis = new FileInputStream(inputFile);
//	            XSSFWorkbook workbook = new XSSFWorkbook(fis);
//	            fis.close();
//
//	            // Get the first sheet
//	            sheet = workbook.getSheetAt(0);
//
//	            // Determine the total number of rows with data
//	            int totalRowsWithData = sheet.getLastRowNum() + 1;
//
//	            // Calculate the end row for copying
//	            int endRow = totalRowsWithData * 24;
//
//	            // Copy the data until the end row
//	            int rowCount = 0;
//	            while (rowCount < endRow) {
//	                // Copy rows from 1 to totalRowsWithData
//	                for (int i = 1; i <= totalRowsWithData; i++) {
//	                    XSSFRow originalRow = sheet.getRow(i - 1); // -1 because rows are 0-indexed
//	                    XSSFRow newRow = sheet.createRow(rowCount++);
//	                    if (originalRow != null) {
//	                        for (int j = originalRow.getFirstCellNum(); j < originalRow.getLastCellNum(); j++) {
//	                            XSSFCell originalCell = originalRow.getCell(j);
//	                            XSSFCell newCell = newRow.createCell(j);
//	                            if (originalCell != null) {
//	                            	switch (originalCell.getCellType()) {
//                                    case STRING:
//                                        newCell.setCellValue(originalCell.getStringCellValue());
//                                        break;
//                                    case NUMERIC:
//                                        newCell.setCellValue(originalCell.getNumericCellValue());
//                                    default:
//                                        newCell.setCellValue(originalCell.toString());
//                                }
//	                            }
//	                        }
//	                    }
//	                }
//	            }
//
//	            // Write the output to a new Excel file
//	            FileOutputStream fos = new FileOutputStream(outputFile);
//	            workbook.write(fos);
//	            fos.close();
//	            workbook.close();
//	            System.out.println("Data copied successfully!");
//
//	        } catch (Exception e) {
//	            e.printStackTrace();
//	        }
//	 }
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
			 GetUserData.getCredentials(USER_DATA_FILE_PATH);
			 writeUserData();
//			 ExcelDataCopier();
			 excelToCsv();	
			 System.out.print("Your data is in the file path - C:\\apache-jmeter-5.5\\apache-jmeter-5.5\\bin\\TestData.csv");
		 }
		 catch(Exception e){
			 e.printStackTrace();
		 }
	 }
}
