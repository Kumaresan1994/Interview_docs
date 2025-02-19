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

public class XMLComparisonUtil {

    public static void compareXMLAndGenerateExcel(String expectedFilePath, String actualFilePath, String excelOutputPath) throws IOException {
        // Read XML files
        String expectedXml = new String(Files.readAllBytes(Paths.get(expectedFilePath)));
        String actualXml = new String(Files.readAllBytes(Paths.get(actualFilePath)));

        // Compare XML files
        Diff diff = DiffBuilder.compare(Input.fromString(expectedXml))
                .withTest(Input.fromString(actualXml))
                .ignoreWhitespace()
                .checkForSimilar()
                .build();

        // Collect differences
        List<String[]> differencesList = new ArrayList<>();
        differencesList.add(new String[]{"Line", "XPath", "Expected Value", "Actual Value"}); // Header row

        if (diff.hasDifferences()) {
            for (Difference d : diff.getDifferences()) {
                String xPath = d.getComparison().getControlDetails().getXPath();
                String expectedValue = d.getComparison().getControlDetails().getValue() != null 
                        ? d.getComparison().getControlDetails().getValue().toString() 
                        : "N/A";
                String actualValue = d.getComparison().getTestDetails().getValue() != null 
                        ? d.getComparison().getTestDetails().getValue().toString() 
                        : "N/A";

                int lineNumber = findLineNumber(expectedXml, expectedValue);

                // Add to list
                differencesList.add(new String[]{String.valueOf(lineNumber), xPath, expectedValue, actualValue});
            }
        }

        // Write differences to Excel
        writeDifferencesToExcel(excelOutputPath, differencesList);
    }

    private static int findLineNumber(String content, String value) {
        if (value.equals("N/A")) return -1;
        String[] lines = content.split("\n");
        for (int i = 0; i < lines.length; i++) {
            if (lines[i].contains(value)) {
                return i + 1;
            }
        }
        return -1;
    }

    private static void writeDifferencesToExcel(String fileName, List<String[]> data) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("XML Differences");

        // Header style
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        // Highlight style
        CellStyle highlightStyle = workbook.createCellStyle();
        highlightStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Write data
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < data.get(i).length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(data.get(i)[j]);

                // Apply styles
                if (i == 0) cell.setCellStyle(headerStyle);
                else cell.setCellStyle(highlightStyle);
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
