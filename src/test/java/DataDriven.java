import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

    public ArrayList<String> getData(String testcaseName) throws IOException {
        ArrayList<String> testData = new ArrayList<>();

        String filePath = "C://Users//Tierless_soul//eclipse-workspace//ExcelDriven//demodata.xlsx";
        try (FileInputStream fis = new FileInputStream(filePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            int sheetCount = workbook.getNumberOfSheets();

            for (int i = 0; i < sheetCount; i++) {
                if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
                    XSSFSheet sheet = workbook.getSheetAt(i);

                    // Identify "TestCases" column index
                    Iterator<Row> rows = sheet.iterator();
                    Row headerRow = rows.next();
                    int testCaseColumnIndex = getColumnIndex(headerRow, "TestCases");

                    // Iterate rows to find the matching test case row
                    while (rows.hasNext()) {
                        Row row = rows.next();
                        Cell cell = row.getCell(testCaseColumnIndex);

                        if (cell != null && cell.getStringCellValue().equalsIgnoreCase(testcaseName)) {
                            // Add all cell values from the matching row
                            for (Cell c : row) {
                                testData.add(c.getStringCellValue());
                            }
                            break; // Exit after finding the matching row
                        }
                    }
                    break; // Exit after processing the correct sheet
                }
            }
        }

        return testData;
    }

    private int getColumnIndex(Row headerRow, String columnName) {
        int index = 0;
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return index;
            }
            index++;
        }
        throw new RuntimeException("Column '" + columnName + "' not found in the header row.");
    }

    public static void main(String[] args) {
        DataDriven dd = new DataDriven();
        try {
            ArrayList<String> data = dd.getData("Purchase");
            System.out.println("Data for 'Purchase' test case: " + data);
        } catch (IOException e) {
            System.err.println("Failed to read test data: " + e.getMessage());
        }
    }
}
