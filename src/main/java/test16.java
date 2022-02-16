import java.io.FileInputStream;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class test16 {
	public static void main(String[] args) throws Throwable {
		HashMap<String, Integer> results = new HashMap<String, Integer>();
		Workbook workbook = WorkbookFactory.create(new FileInputStream(args[0]));
		Sheet sheet = workbook.getSheet(args[1]);
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			String A = getCellValue(row.getCell(0));
			String B = getCellValue(row.getCell(1));
			String M = getCellValue(row.getCell(12));
			if (A == null) {
				continue;
			}
			if (B == null) {
				continue;
			}
			try {
				int n = Integer.valueOf(A);
				if (n <= 35) {
					continue;
				}
			} catch (NumberFormatException e) {
				continue;
			}
			M = M == null ? "" : M.trim();
			M = M.replace("　", "");
			if (results.containsKey(M)) {
				results.put(M, results.get(M) + 1);
			} else {
				results.put(M, 1);
			}
			System.out.println(A + " " + B + " " + M);
		}
		System.out.println("---");
		System.out.println(results.get(""));
		System.out.println(results.get("完了"));
		System.out.println(results.get("") + results.get("完了"));
	}

	static String getCellValue(Cell cell) {
		if (cell == null) {
			return null;
		}
		return new DataFormatter().formatCellValue(cell);
	}
}
