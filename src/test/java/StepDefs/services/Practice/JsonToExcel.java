package StepDefs.services.Practice;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class JsonToExcel {
    public static void main(String[] args) {
        // Path to the JSON file
        String jsonFilePath = ".//resources//files//Sample_json.json"; // Replace with your file path

        // Path to the output Excel file
        String excelFilePath = ".//resources//files//Sample_output.xlsx"; // Replace with your desired output path

        try {
            // Read the JSON file content as a String
            String jsonContent = new String(Files.readAllBytes(Paths.get(jsonFilePath)));

            // Check if the JSON content starts with '{' or '['
            if (jsonContent.trim().startsWith("[")) {
                // Parse it as a JSONArray
                JSONArray jsonArray = new JSONArray(jsonContent);
                writeToExcel(jsonArray, excelFilePath);
            } else if (jsonContent.trim().startsWith("{")) {
                // Parse it as a JSONObject
                JSONObject jsonObject = new JSONObject(jsonContent);

                // If it's a JSONObject, wrap it in a JSONArray for consistent processing
                JSONArray jsonArray = new JSONArray();
                jsonArray.put(jsonObject);
                writeToExcel(jsonArray, excelFilePath);
            } else {
                System.out.println("Invalid JSON format. It must start with '{' or '['.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Method to write JSONArray to Excel
    private static void writeToExcel(JSONArray jsonArray, String excelFilePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("JSON Data");

        // Write headers based on the first object in the array
        Row headerRow = sheet.createRow(0);
        JSONObject firstObject = jsonArray.getJSONObject(0);
        int headerIndex = 0;
        for (String key : firstObject.keySet()) {
            Cell cell = headerRow.createCell(headerIndex++);
            cell.setCellValue(key);
        }

        // Write data rows for each JSON object in the array
        int rowIndex = 1;
        for (int i = 0; i < jsonArray.length(); i++) {
            JSONObject jsonObject = jsonArray.getJSONObject(i);
            Row row = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            for (String key : jsonObject.keySet()) {
                Cell cell = row.createCell(cellIndex++);
                cell.setCellValue(jsonObject.get(key).toString());
            }
        }

        // Write the Excel file to disk
        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
        }

        workbook.close();
        System.out.println("Excel file created successfully!");
    }
}
