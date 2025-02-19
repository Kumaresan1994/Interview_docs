import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xmlunit.builder.DiffBuilder;
import org.xmlunit.builder.Input;
import org.xmlunit.diff.Diff;
import org.xmlunit.diff.Difference;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class XMLComparisonToExcel {
    public static void main(String[] args) {
        try {
            // Read XML files
            String expectedXml = new String(Files.readAllBytes(Paths.get("expected.xml")));
            String actualXml = new String(Files.readAllBytes(Paths.get("actual.xml")));

            // Compare XML files
            Diff diff = DiffBuilder.compare(Input.fromString(expectedXml))
                    .withTest(Input.fromString(actualXml))
                    .ignoreWhitespace()
                    .checkForSimilar() // Use checkForIdentical() for strict match
                    .build();

            // Collect differences
            List<String[]> differencesList = new ArrayList<>();
            differencesList.add(new String[]{"Line", "XPath", "Expected Value", "Actual Value"}); // Header row

            if (diff.hasDifferences()) {
                for (Difference d : diff.getDifferences()) {
                    String description = d.toString();

                    // Extract values using regex
                    String expectedValue = extractValue(description, "Expected '(.*?)' but");
                    String actualValue = extractValue(description, "but was '(.*?)' - at");
                    String xPath = extractValue(description, "- at (.*?)$");
                    int lineNumber = findLineNumber(expectedXml, expectedValue);

                    // Add to differences list
                    differencesList.add(new String[]{String.valueOf(lineNumber), xPath, expectedValue, actualValue});
                }
            }

            // Write differences to Excel file
            writeDifferencesToExcel("differences.xlsx", differencesList);
            System.out.println("Differences written to differences.xlsx");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String extractValue(String text, String regex) {
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(text);
        return matcher.find() ? matcher.group(1) : "N/A";
    }

    private static int findLineNumber(String content, String value) {
        if (value.equals("N/A")) return -1;
        String[] lines = content.split("\n");
        for (int i = 0; i < lines.length; i++) {
            if (lines[i].contains(value)) {
                return i + 1; // Line numbers start from 1
            }
        }
        return -1;
    }

    private static void writeDifferencesToExcel(String fileName, List<String[]> data) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("XML Differences");

        // Create header style
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        // Create highlight style for differences
        CellStyle highlightStyle = workbook.createCellStyle();
        highlightStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Write data to sheet
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < data.get(i).length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(data.get(i)[j]);

                // Apply styles
                if (i == 0) {
                    cell.setCellStyle(headerStyle); // Header styling
                } else {
                    cell.setCellStyle(highlightStyle); // Highlight differences
                }
            }
        }

        // Auto-size columns
        for (int i = 0; i < data.get(0).length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write to file
        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
            workbook.write(outputStream);
        }

        workbook.close();
    }
}
