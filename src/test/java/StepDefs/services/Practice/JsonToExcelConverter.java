package StepDefs.services.Practice;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.FileOutputStream;
import java.io.IOException;

public class JsonToExcelConverter {

    public static void main(String[] args) {
        String jsonString = "[" +
                "{\"code\":\"541512\",\"description\":Computer Systems Design Services,\"typeDescription\":North American Industry Classification System 2022,\"typeDnBCode\":3599,\"priority\":1}," +
                "{\"code\":\"73790000\",\"description\":Computer related services,\"typeDescription\":D&B Standard Industry Code,\"typeDnBCode\":3599,\"priority\":1}," +
                "{\"code\":\"6203\",\"description\":Computer facilities management activities,\"typeDescription\":NACE Revision 2,\"typeDnBCode\":29104,\"priority\":1}," +
                "{\"code\":\"7379\",\"description\":Computer related services,\"typeDescription\":US Standard Industry Code 1987 - 4 digit,\"typeDnBCode\":399,\"priority\":1}," +
                "{\"code\":\"42\",\"description\":Computer System Design Services,\"typeDescription\":D&B Hoovers Industry Classification,\"typeDnBCode\":35912,\"priority\":1}," +
                "{\"code\":\"I\",\"description\":Services,\"typeDescription\":D&B Standard Major Industry Code,\"typeDnBCode\":24657,\"priority\":1}," +
                "{\"code\":\"6202\",\"description\":Computer consultancy and computer facilities management activities,\"typeDescription\":ISIC Revision 4,\"typeDnBCode\":42726,\"priority\":1}," +
                "]";
        convertJsonToExcel(jsonString, ".//resources//files//Sample_output.xlsx");
    }

    public static void convertJsonToExcel(String jsonString, String excelFilePath) {
        // Create a Workbook
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("code");
        headerRow.createCell(1).setCellValue("description");
        headerRow.createCell(2).setCellValue("typeDescription");
        headerRow.createCell(3).setCellValue("typeDnBCode");
        headerRow.createCell(4).setCellValue("priority");

        // Parse JSON
        JSONArray jsonArray = new JSONArray(jsonString);

        // Fill the Excel sheet with JSON data
        for (int i = 0; i < jsonArray.length(); i++) {
            JSONObject jsonObject = jsonArray.getJSONObject(i);
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(jsonObject.getString("code"));
            row.createCell(1).setCellValue(jsonObject.getString("description"));
            row.createCell(2).setCellValue(jsonObject.getString("typeDescription"));
            row.createCell(3).setCellValue(jsonObject.getInt("typeDnBCode"));
            row.createCell(4).setCellValue(jsonObject.getInt("priority"));
        }

        // Write to Excel file
        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
            System.out.println("JSON data is translated to Excel file successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
