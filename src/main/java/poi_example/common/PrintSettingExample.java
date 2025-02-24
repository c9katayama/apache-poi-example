package poi_example.common;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excelファイルの印刷設定を行うサンプル。
 */
public class PrintSettingExample {
    public static void main(String[] args) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            // シートの作成
            XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Print Settings");

            // 印刷設定
            PrintSetup printSetup = sheet.getPrintSetup();
            printSetup.setPaperSize(PrintSetup.A4_PAPERSIZE); // 用紙サイズをA4に設定
            printSetup.setLandscape(false); // 印刷の向きを縦に設定 (false: 縦, true: 横)
            printSetup.setFitWidth((short) 1); // 横幅を1ページに収める
            printSetup.setFitHeight((short) 0); // 縦幅はページ数に合わせて調整

            // ヘッダーとフッターの設定
            sheet.getHeader().setCenter("&F");  // ファイル名を中央に
            sheet.getHeader().setLeft("confidential");  // 左に機密表示
            sheet.getHeader().setRight("Page &P of &N");  // 右にページ番号

            sheet.getFooter().setLeft("&D &T");  // 左に日付と時刻
            sheet.getFooter().setCenter("Copyright (c) 2024");  // 中央に著作権
            sheet.getFooter().setRight("Prepared by ...");  // 右に作成者

            // マージンの設定 (単位: インチ)
            sheet.setMargin(PageMargin.FOOTER, 0.5);
            sheet.setMargin(PageMargin.HEADER, 0.5);
            sheet.setMargin(PageMargin.LEFT, 0.75);
            sheet.setMargin(PageMargin.RIGHT, 0.75);
            sheet.setMargin(PageMargin.TOP, 1.0);
            sheet.setMargin(PageMargin.BOTTOM, 1.0);

            // ファイルに保存
            new File("output").mkdirs();
            try (FileOutputStream fileOut = new FileOutputStream("output/print_setting_example.xlsx")) {
                workbook.write(fileOut);
            }
        }
        System.out.println("印刷設定済みのExcelファイルを作成しました。");
    }
}