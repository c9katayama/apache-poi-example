package poi_example.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * 大量のデータを効率的に出力するためのSXSSFWorkbookを使用したサンプル。
 * メモリ内に保持する行数を制限し、一定以上の行は一時ファイルに書き出します。
 */
public class LargeWorkbookExportExample {
    public static void main(String[] args) throws IOException {
        // メモリ内に保持する行数を100行に設定してSXSSFWorkbookを作成
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(100)) {
            Sheet sheet = workbook.createSheet("Large Data");

            // ヘッダー行の作成
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("ID");
            headerRow.createCell(1).setCellValue("名前");
            headerRow.createCell(2).setCellValue("データ");

            // 大量のデータを出力（例として100000行）
            for (int i = 1; i <= 100000; i++) {
                Row row = sheet.createRow(i);

                Cell idCell = row.createCell(0);
                idCell.setCellValue(i);

                Cell nameCell = row.createCell(1);
                nameCell.setCellValue("名前" + i);

                Cell dataCell = row.createCell(2);
                dataCell.setCellValue("データ" + i);

                // 進捗状況を1000行ごとに表示
                if (i % 1000 == 0) {
                    System.out.println(i + "行目まで処理しました。");
                }
            }

            // ファイルに保存
            new File("output").mkdirs();
            String outputFilePath = "output/large_data.xlsx";
            try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                workbook.write(fileOut);
                System.out.println("大量データを含むExcelファイルを作成しました。ファイルパス: " + outputFilePath);
            }

            // 一時ファイルを削除
            workbook.dispose();
        }
    }
}