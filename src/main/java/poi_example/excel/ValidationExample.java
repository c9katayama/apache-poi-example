package poi_example.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excelファイルに入力規則を設定するサンプル。
 */
public class ValidationExample {
    public static void main(String[] args) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Validation");

            // リストによるデータ検証
            DataValidationHelper validationHelper = sheet.getDataValidationHelper();

            // リストの選択肢を別のセルに設定
            Row row0 = sheet.createRow(0);
            row0.createCell(2).setCellValue("赤");
            row0.createCell(3).setCellValue("青");
            row0.createCell(4).setCellValue("緑");

            // リストによる入力規則を設定
            CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
            DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(
                    new String[] { "赤", "青", "緑" });
            DataValidation dataValidation = validationHelper.createValidation(constraint, addressList);

            // エラーメッセージの設定
            dataValidation.setShowErrorBox(true);
            dataValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
            dataValidation.createErrorBox("エラー", "リストから値を選択してください。");

            // 入力メッセージの設定
            dataValidation.setShowPromptBox(true);
            dataValidation.createPromptBox("選択", "色を選択してください。");

            sheet.addValidationData(dataValidation);

            // 数値の範囲による入力規則
            CellRangeAddressList numberRange = new CellRangeAddressList(1, 1, 0, 0);
            DataValidationConstraint numberConstraint = validationHelper.createIntegerConstraint(
                    DataValidationConstraint.OperatorType.BETWEEN, "1", "100");
            DataValidation numberValidation = validationHelper.createValidation(
                    numberConstraint, numberRange);

            numberValidation.setShowErrorBox(true);
            numberValidation.setErrorStyle(DataValidation.ErrorStyle.STOP);
            numberValidation.createErrorBox("エラー", "1から100の間の数値を入力してください。");

            sheet.addValidationData(numberValidation);

            // ファイルに保存
            new File("output").mkdirs();
            try (FileOutputStream fileOut = new FileOutputStream("output/validation_example.xlsx")) {
                workbook.write(fileOut);
            }
        }
        System.out.println("入力規則を設定したExcelファイルを作成しました。");
    }
}