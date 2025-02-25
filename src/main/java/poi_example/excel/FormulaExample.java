package poi_example.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excelファイルに数式を設定するサンプル。
 */
public class FormulaExample {
    public static void main(String[] args) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Formula Sheet");

            // データ入力
            Row row1 = sheet.createRow(0);
            row1.createCell(0).setCellValue(10);
            row1.createCell(1).setCellValue(20);

            // 数式の設定
            Row row2 = sheet.createRow(1);
            Cell formulaCell = row2.createCell(0);
            formulaCell.setCellFormula("SUM(A1:B1)");

            // 数式の評価
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateFormulaCell(formulaCell);

            // ファイルに保存
            new File("output").mkdirs();
            String outputFilePath = "output/formula.xlsx";
            try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                workbook.write(fileOut);
                System.out.println("数式を含むExcelファイルを作成しました。ファイルパス: " + outputFilePath);
            }
        }
    }
}