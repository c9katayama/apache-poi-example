package poi_example.common;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EncryptionExample {
    public static void main(String[] args) throws IOException, InvalidFormatException, Exception {
        // 新規ワークブック作成
        try (Workbook workbook = new XSSFWorkbook()) {
            // テスト用のwrokbook作成

            Sheet sheet = workbook.createSheet("Sheet1");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Hello");

            // ファイル出力ストリーム
            File file = new File("output/encrypted_workbook.xlsx");
            file.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(file);
                    POIFSFileSystem fs = new POIFSFileSystem()) {
                EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
                Encryptor enc = info.getEncryptor();
                enc.confirmPassword("password");

                // ワークブックを暗号化して書き込み
                try (OutputStream os = enc.getDataStream(fs)) {
                    workbook.write(os);
                }
                fs.writeFilesystem(fos);
            }
        }
        System.out.println("暗号化されたExcelファイルを作成しました。");

        // 復号化 (サンプルとして読み込みのみ)
        try (FileInputStream fis = new FileInputStream("output/encrypted_workbook.xlsx")) {
            Workbook workbook = WorkbookFactory.create(fis, "password");
            System.out.println("Excelファイルを復号化して読み込みました。" + workbook.getSheetName(0));
            // ... (ワークブックの操作) ...
        } catch (Exception e) {
            System.err.println("復号化に失敗しました: " + e.getMessage());
        }
    }
}