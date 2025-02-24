package poi_example.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

/**
 * 大量のデータを含むExcelファイルを効率的に読み込むサンプル。
 * イベントモデルを使用して、メモリ使用量を抑えながら処理を行います。
 */
public class LargeWorkbookReaderExample {
    public static void main(String[] args) throws Exception {
        // 大量のデータを含むExcelファイルを作成
        LargeWorkbookExportExample.main(args);
        File file = new File("output/large_data.xlsx");
        try (InputStream is = new FileInputStream(file);
                OPCPackage pkg = OPCPackage.open(is)) {

            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings sst = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();

            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles,
                    sst,
                    new SheetContentsHandler() {
                        private int currentRow = -1;
                        private int processedRows = 0;

                        @Override
                        public void startRow(int rowNum) {
                            currentRow = rowNum;
                        }

                        @Override
                        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                            // セルの値を処理
                            // 実際の処理ではここでデータベースへの保存や集計などを行う
                            System.out.print(formattedValue + "\t");
                        }

                        @Override
                        public void endRow(int rowNum) {
                            System.out.println(); // 行の終わりで改行
                            processedRows++;

                            // 1000行ごとに進捗を表示
                            if (processedRows % 1000 == 0) {
                                System.out.println(processedRows + "行を処理しました。現在の行：" + currentRow);
                            }
                        }

                        @Override
                        public void headerFooter(String text, boolean isHeader, String tagName) {
                            // ヘッダーやフッターの処理が必要な場合はここに実装
                        }
                    },
                    new DataFormatter(),
                    false // フォーミュラの評価を行わない
            );
            
            // XMLReaderの設定
            XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
            parser.setContentHandler(handler);

            // シートの処理
            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (sheets.hasNext()) {
                try (InputStream sheetStream = sheets.next()) {
                    System.out.println("シート名: " + sheets.getSheetName());
                    InputSource sheetSource = new InputSource(sheetStream);
                    parser.parse(sheetSource);
                }
            }
        }

        System.out.println("ファイルの読み込みが完了しました。");
    }
}