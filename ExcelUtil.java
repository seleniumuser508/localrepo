package framework;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	
	public static Workbook workbook;
    public static Sheet sheet;
    public static Cell cell;
    
    public void getExcelFile(String excelFilePath) throws FileNotFoundException, IOException {
        FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
        workbook = new XSSFWorkbook(excelFile);
    }
    
    public void getSheet(String sheetName) {
        sheet = workbook.getSheet(sheetName);
    }
    
    public int getRowCount() {
        Iterator<Row> iterator = sheet.iterator();
        int rowCount = 0;
        while (iterator.hasNext()) {
            rowCount++;
            iterator.next();
        }
        return rowCount;
    }
    
    public static int getColCount() {
        return sheet.getRow(0).getPhysicalNumberOfCells();
    }
    
    public String getCellValue(int row, int col) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Iterator<Row> iterator = sheet.iterator();
        Row nextRow = null;

        for (int i = 1; i <= row; i++) {
            nextRow = iterator.next();
        }

        cell = nextRow.getCell(col);
        CellValue cellValue = evaluator.evaluate(cell);

        try {
            return cellValue.getStringValue();
        } catch (Exception e) {
            return "";
        }

    }
    
    public static void closeObjects() {
        workbook = null;
        sheet = null;
        cell = null;
    }

}
