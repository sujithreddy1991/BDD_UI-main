package StepDefs.services.Practice;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashSet;
import java.util.Set;

public class JsonToExcel {
    public static void main(String[] args) {
        // Paths for input and output
        String jsonFilePath = ".//resources//files//Sample_json.json";
        String excelFilePath = ".//resources//files//Sample_output.xlsx";

        try {
            String jsonContent = new String(Files.readAllBytes(Paths.get(jsonFilePath)));
            JSONArray jsonArray;

            // Determine if the content is a JSON array or object
            if (jsonContent.trim().startsWith("[")) {
                jsonArray = new JSONArray(jsonContent);
            } else if (jsonContent.trim().startsWith("{")) {
                JSONObject jsonObject = new JSONObject(jsonContent);
                jsonArray = new JSONArray();
                jsonArray.put(jsonObject);
            } else {
                throw new IllegalArgumentException("Invalid JSON format. It must start with '{' or '['.");
            }

            writeToExcel(jsonArray, excelFilePath);

        } catch (IOException | IllegalArgumentException e) {
            System.err.println("Error processing the file: " + e.getMessage());
        }
    }

    private static void writeToExcel(JSONArray jsonArray, String excelFilePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("JSON Data");

        // Collect unique headers
        Set<String> headers = new HashSet<>();
        for (int i = 0; i < jsonArray.length(); i++) {
            JSONObject jsonObject = jsonArray.getJSONObject(i);
            headers.addAll(jsonObject.keySet());
        }

        // Create header row
        Row headerRow = sheet.createRow(0);
        int headerIndex = 0;
        for (String header : headers) {
            headerRow.createCell(headerIndex++).setCellValue(header);
        }

        // Populate rows with data
        for (int rowIndex = 1; rowIndex <= jsonArray.length(); rowIndex++) {
            JSONObject jsonObject = jsonArray.getJSONObject(rowIndex - 1);
            Row row = sheet.createRow(rowIndex);
            int cellIndex = 0;
            for (String header : headers) {
                Cell cell = row.createCell(cellIndex++);
                cell.setCellValue(jsonObject.optString(header, ""));
            }
        }

        // Write to Excel file
        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
        } finally {
            workbook.close();
        }

        System.out.println("Excel file created successfully!");
    }
}
