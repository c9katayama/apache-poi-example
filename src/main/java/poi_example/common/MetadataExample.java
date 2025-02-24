package poi_example.common;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excelファイルのメタデータを設定するサンプル。
 */
public class MetadataExample {
    public static void main(String[] args) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            // ドキュメントプロパティ (コアプロパティ) の設定
            POIXMLProperties props = workbook.getProperties();
            props.getCoreProperties().setCreator("Author Name");
            props.getCoreProperties().setTitle("Sample Excel File");
            props.getCoreProperties().setCategory("メタデータ設定のサンプル");
            props.getCoreProperties().setDescription("これはPOIによるメタデータ設定のサンプルファイルです。");
            props.getCoreProperties().setLastModifiedByUser("最終更新者名");
            props.getCoreProperties().setCreated(java.util.Optional.of(new java.util.Date()));
            props.getCoreProperties().setModified(java.util.Optional.of(new java.util.Date()));

            // ファイルに保存
            new File("output").mkdirs();
            try (FileOutputStream fileOut = new FileOutputStream("output/metadata_example.xlsx")) {
                workbook.write(fileOut);
            }
        }
        System.out.println("メタデータ設定済みのExcelファイルを作成しました。");
    }
}