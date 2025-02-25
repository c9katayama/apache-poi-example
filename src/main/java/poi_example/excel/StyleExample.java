package poi_example.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excelファイルにフォントサイズやセルの色などのスタイルを適用するサンプル。
 */
public class StyleExample {
    public static void main(String[] args) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Styles");

            // フォントスタイルの作成
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 14);
            headerFont.setColor(IndexedColors.RED.getIndex());

            // セルスタイルの作成
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setBorderBottom(BorderStyle.THICK);
            headerStyle.setBorderTop(BorderStyle.THICK);
            headerStyle.setBorderLeft(BorderStyle.THICK);
            headerStyle.setBorderRight(BorderStyle.THICK);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);

            // ヘッダー行の作成
            Row headerRow = sheet.createRow(0);
            Cell headerCell = headerRow.createCell(0);
            headerCell.setCellValue("スタイル適用例");
            headerCell.setCellStyle(headerStyle);

            // 列幅の自動調整
            sheet.autoSizeColumn(0);

            // ファイルに保存
            new File("output").mkdirs();
            String outputFilePath = "output/style_example.xlsx";
            try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                workbook.write(fileOut);
                System.out.println("スタイルを適用したExcelファイルを作成しました。ファイルパス: " + outputFilePath);
            }
        }
    }
}