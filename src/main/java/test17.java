import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

public class test17 {

	public static void main(String[] args) throws Throwable {
		Workbook workbook1 = WorkbookFactory.create(new FileInputStream(args[0]));
		Workbook workbook2 = WorkbookFactory.create(new FileInputStream(args[1]));

		List<String> sheetNames = new ArrayList<String>();
		addAllSheetNames(sheetNames, workbook1);
		addAllSheetNames(sheetNames, workbook2);

		for (String sheetName : sheetNames) {
			Sheet sheet1 = workbook1.getSheet(sheetName);
			Sheet sheet2 = workbook2.getSheet(sheetName);

			if (sheet1 == null) {
				System.out.println(sheetName + "\tシート追加");
				continue;
			}
			if (sheet2 == null) {
				System.out.println(sheetName + "\tシート削除");
				continue;
			}

			for (int i = 0; i <= sheet1.getLastRowNum() || i <= sheet2.getLastRowNum(); i++) {
				Row row1 = sheet1.getRow(i);
				Row row2 = sheet2.getRow(i);
				short lastCellNum = 0;
				if (row1 != null) {
					lastCellNum = row1.getLastCellNum();
				}
				if (row2 != null && row2.getLastCellNum() > lastCellNum) {
					lastCellNum = row2.getLastCellNum();
				}
				for (int s = 0; s <= lastCellNum; s++) {
					String value1 = row1 == null ? null : getCellValue(row1.getCell(s));
					String value2 = row2 == null ? null : getCellValue(row2.getCell(s));
					if (value1 == null && value2 == null) {
						continue;
					}
					if (value1 != null && value1.equals(value2)) {
						continue;
					}
					String cellLocation = new CellReference(i, s).formatAsString();
					System.out.println(sheetName + "\t"
							+ cellLocation + "\t\""
							+ value1.replace("\"", "\"\"") + "\"\t\""
							+ value2.replace("\"", "\"\"") + "\"");
				}
			}
		}

	}

	static void addAllSheetNames(List<String> names, Workbook workbook) {
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			if (!names.contains(workbook.getSheetName(i))) {
				names.add(workbook.getSheetName(i));
			}
		}
	}

	static String getCellValue(Cell cell) {
		if (cell == null) {
			return null;
		}
		return new DataFormatter().formatCellValue(cell);
	}
}
