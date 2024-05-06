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

import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;

import org.json.JSONObject;

public class GetUserData {
	 	static Map<String, String> UserData = new HashMap<>();;
	 	static Map<String, String[]> accountData = new HashMap<>();
	 	static String [] username = new String[1801];
	    static int RowNum = 1;
	    static FileOutputStream fileOut;
	    static XSSFSheet sheet;
	    static String access_Token;
	    
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
            workbook.close();
            fileIn.close();
        }
	}
	public static void getAccessToken() throws IOException {
		OkHttpClient client = new OkHttpClient().newBuilder()
				  .build();
				MediaType mediaType = MediaType.parse("text/plain");
				RequestBody body = RequestBody.create(mediaType, "");
				Request request = new Request.Builder()
				  .url("https://cona--lfln012.sandbox.my.salesforce.com/services/oauth2/token?client_id=3MVG9EMJF5MdlzDpwkTqS3Z06ccHdrpzYbwlj.PSZWX4DESHHs.D4xsv6vye4JPLDl9UpdIlpnzNfUx500vQq&client_secret=64C750BE75DE190D1FC3A6FC95E82BE7BBF96529356F7504BA8BE7B29DFB0B26&grant_type=client_credentials")
				  .method("POST", body)
				  .addHeader("Cookie", "BrowserId=TNnxFZQrEe65CA1jhtBqdQ; CookieConsentPolicy=0:1; LSKey-c$CookieConsentPolicy=0:1")
				  .build();
				Response response = client.newCall(request).execute();
				// Parse JSON response
				String responseBody = response.body().string();
		        JSONObject json = new JSONObject(responseBody);
		        // Get the access token
		        access_Token = json.getString("access_token");
				response.close();
	}
}

