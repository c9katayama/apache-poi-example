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
 * ExcelファイルにVLOOKUPを使用した数式を設定するサンプル。
 */
public class FormulaVLookupExample {
    public static void main(String[] args) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("VLOOKUP Example");

            // 参照テーブルの作成（商品コードと価格）
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("商品コード");
            headerRow.createCell(1).setCellValue("価格");

            Row dataRow1 = sheet.createRow(1);
            dataRow1.createCell(0).setCellValue("A001");
            dataRow1.createCell(1).setCellValue(1000);

            Row dataRow2 = sheet.createRow(2);
            dataRow2.createCell(0).setCellValue("A002");
            dataRow2.createCell(1).setCellValue(2000);

            Row dataRow3 = sheet.createRow(3);
            dataRow3.createCell(0).setCellValue("A003");
            dataRow3.createCell(1).setCellValue(3000);

            // 検索値の入力
            Row searchRow = sheet.createRow(5);
            searchRow.createCell(0).setCellValue("検索する商品コード：");
            searchRow.createCell(1).setCellValue("A002");

            // VLOOKUPの数式を設定
            Row resultRow = sheet.createRow(6);
            resultRow.createCell(0).setCellValue("価格：");
            Cell formulaCell = resultRow.createCell(1);
            formulaCell.setCellFormula("VLOOKUP(B6,$A$1:$B$4,2,FALSE)");

            // 数式の評価
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateFormulaCell(formulaCell);

            // 列幅の自動調整
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);

            // ファイルに保存
            new File("output").mkdirs();
            try (FileOutputStream fileOut = new FileOutputStream("output/vlookup_example.xlsx")) {
                workbook.write(fileOut);
            }
        }
        System.out.println("VLOOKUPを含むExcelファイルを作成しました。");
    }
}