package poi_example.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConditionalFormattingExample {
	public static void main(String[] args) throws IOException {
		try (Workbook workbook = new XSSFWorkbook()) {
			Sheet sheet = workbook.createSheet("条件付き書式サンプル");

			// サンプルデータの作成
			Object[][] data = { { 10, 55, 15, 80, 20 }, { 30, 40, 15, 45, 50 }, { 15, 25, 30, 15, 60 } };

			// データの入力
			for (int i = 0; i < data.length; i++) {
				Row row = sheet.createRow(i);
				for (int j = 0; j < data[i].length; j++) {
					Cell cell = row.createCell(j);
					cell.setCellValue((Integer) data[i][j]);
				}
			}

			// 条件付き書式の設定
			SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

			// 1. 50以上の値を赤色で表示
			ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "50");
			PatternFormatting pattern1 = rule1.createPatternFormatting();
			pattern1.setFillBackgroundColor(IndexedColors.RED.getIndex());
			pattern1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

			// 2. 20未満の値を緑色で表示
			ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, "20");
			PatternFormatting pattern2 = rule2.createPatternFormatting();
			pattern2.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
			pattern2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

			// 3. 重複値を黄色でハイライト
			ConditionalFormattingRule rule3 = sheetCF.createConditionalFormattingRule("COUNTIF($A$1:$E$3,A1)>1");
			PatternFormatting pattern3 = rule3.createPatternFormatting();
			pattern3.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
			pattern3.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

			// 条件付き書式を適用する範囲の設定
			CellRangeAddress[] regions = { CellRangeAddress.valueOf("A1:E3") };

			// ルールの適用
			sheetCF.addConditionalFormatting(regions, rule1);
			sheetCF.addConditionalFormatting(regions, rule2);
			sheetCF.addConditionalFormatting(regions, rule3);

			// 列幅の自動調整
			for (int i = 0; i < 5; i++) {
				sheet.autoSizeColumn(i);
			}

			// ファイルの保存
			new File("output").mkdirs();
			String outputFilePath = "output/conditional_formatting_example.xlsx";
			try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
				workbook.write(fileOut);
				System.out.println("条件付き書式を含むExcelファイルを作成しました。ファイルパス: " + outputFilePath);
			}
		}
	}
}