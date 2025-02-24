package poi_example.word;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * Wordドキュメントに表を作成するサンプル。
 */
public class TableExample {
    public static void main(String[] args) throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {
            // 表の作成（3行2列）
            XWPFTable table = document.createTable(3, 2);

            // 表のスタイル設定
            table.setWidth("100%");
            table.setTableAlignment(TableRowAlign.CENTER);

            // ヘッダー行の設定
            XWPFTableRow headerRow = table.getRow(0);
            headerRow.getCell(0).setText("項目");
            headerRow.getCell(1).setText("値");

            // データ行の設定
            XWPFTableRow row1 = table.getRow(1);
            row1.getCell(0).setText("商品A");
            row1.getCell(1).setText("1,000円");

            XWPFTableRow row2 = table.getRow(2);
            row2.getCell(0).setText("商品B");
            row2.getCell(1).setText("2,000円");

            // セルの書式設定
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    // 段落の配置を中央に
                    XWPFParagraph paragraph = cell.getParagraphs().get(0);
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                }
            }

            // ファイルに保存
            new File("output").mkdirs();
            try (FileOutputStream out = new FileOutputStream("output/table_example.docx")) {
                document.write(out);
            }
        }
        System.out.println("表を含むWordファイルを作成しました。");
    }
}