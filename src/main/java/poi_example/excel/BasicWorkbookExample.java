package poi_example.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excelファイルを作成するサンプル。
 */
public class BasicWorkbookExample {
	public static void main(String[] args) throws IOException {
		// 新規ワークブック作成
		try (Workbook workbook = new XSSFWorkbook()) {
			// シート作成
			Sheet sheet = workbook.createSheet("First Sheet");

			// 行の作成
			Row row = sheet.createRow(0);

			// セルの作成と値の設定
			Cell cell = row.createCell(0);
			cell.setCellValue("Hello Apache POI!");

			// スタイルの設定
			CellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cell.setCellStyle(style);

			// ファイルに保存
			new File("output").mkdirs(); // outputディレクトリを作成
			try (FileOutputStream fileOut = new FileOutputStream("output/workbook.xlsx")) {
				workbook.write(fileOut);
			}
		}
		System.out.println("Excelファイルを作成しました。");
	}
}