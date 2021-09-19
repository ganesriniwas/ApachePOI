package Maven.ApachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readExcel {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(
					new FileInputStream(System.getProperty("user.dir") + "\\input\\TestData.xlsx"));
			Iterator<Sheet> sheets = workbook.sheetIterator();
			String expectedSheetName = "UseCase1";
			String expectedTestCase = "extentReport";
			String expectedData = "Data2";

//			-----------------------Find the right sheet-------------------------------
			int counter = 0;
			int sheetNumber = 0;
			while (sheets.hasNext()) {
				if (sheets.next().getSheetName().equalsIgnoreCase(expectedSheetName)) {
					sheetNumber = counter;
					System.out.println("Sheet = " + sheetNumber);
					break;
				}
				counter++;
			}
			XSSFSheet sheet = workbook.getSheetAt(sheetNumber);

//          -------------------------Find the column number for Data---------------------------------
			Iterator<Row> rows = sheet.iterator();
			Iterator<Cell> cells = rows.next().cellIterator();
			int columnNumber = 0;
			counter = 0;
			while (cells.hasNext()) {
				if (cells.next().getStringCellValue().equalsIgnoreCase(expectedData)) {
					columnNumber = counter;
					System.out.println("Coumn = " + columnNumber);
					break;
				}
				counter++;
			}
//			----------------------Find the row for Test Cases---------------------------
//			Iterator<Row> rows=sheet.iterator();
			int rowNumber = 0;
			counter = 1;
			while (rows.hasNext()) {
				if (rows.next().getCell(0).getStringCellValue().equalsIgnoreCase(expectedTestCase)) {
					rowNumber = counter;
					System.out.println("Row = " + rowNumber);
					break;
				}
				counter++;
			}
			XSSFRow row = sheet.getRow(rowNumber);
//			------------------------------Print Final Result---------------------------
			if ((rowNumber != 0) && (columnNumber != 0)) {
				XSSFCell cell = row.getCell(columnNumber);
				if (cell.getCellType() == CellType.STRING) {
					System.out.println("Expected Value =" + cell.getStringCellValue());
				} else {
					String value = NumberToTextConverter.toText(cell.getNumericCellValue());
					System.out.println("Expected Value =" + value);
				}
			} else {
				System.out.println("Data value not found");
			}
		} catch (FileNotFoundException e) {
			System.out.println("File not available at the location");
		}
	}
}
