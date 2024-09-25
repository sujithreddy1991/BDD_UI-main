package StepDefs.services.Practice;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class JsonToCsv {
    public static void main(String[] args) {
        // Path to the JSON file
        String jsonFilePath = ".//resources//files//Sample_json.json"; // Replace with your file path

        // Path to the output CSV file
        String csvFilePath = ".//resources//files//Sample_outputs.csv"; // Replace with your desired output path

        try {
            // Read JSON file content as a String
            String jsonContent = new String(Files.readAllBytes(Paths.get(jsonFilePath)));

            // Parse JSON content
            JSONArray jsonArray;
            if (jsonContent.trim().startsWith("[")) {
                jsonArray = new JSONArray(jsonContent);
            } else if (jsonContent.trim().startsWith("{")) {
                JSONObject jsonObject = new JSONObject(jsonContent);
                jsonArray = new JSONArray();
                jsonArray.put(jsonObject);
            } else {
                throw new IllegalArgumentException("Invalid JSON format. It must start with '{' or '['.");
            }

            // Write to CSV
            writeToCsv(jsonArray, csvFilePath);
            System.out.println("CSV file created successfully!");

        } catch (IOException | IllegalArgumentException e) {
            System.err.println("Error processing the file: " + e.getMessage());
        }
    }

    private static void writeToCsv(JSONArray jsonArray, String csvFilePath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(csvFilePath))) {
            // Get headers
            JSONObject firstObject = jsonArray.getJSONObject(0);
            StringBuilder header = new StringBuilder();
            for (String key : firstObject.keySet()) {
                header.append(key).append(",");
            }
            // Remove last comma and add a new line
            writer.write(header.substring(0, header.length() - 1) + "\n");

            // Write data
            for (int i = 0; i < jsonArray.length(); i++) {
                JSONObject jsonObject = jsonArray.getJSONObject(i);
                StringBuilder row = new StringBuilder();
                for (String key : firstObject.keySet()) {
                    row.append(jsonObject.optString(key, "")).append(",");
                }
                // Remove last comma and add a new line
                writer.write(row.substring(0, row.length() - 1) + "\n");
            }
        }
    }
}
