package poi_example.excel;

import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 複数の画像をExcelファイルに縦に並べて貼り付けるサンプル。
 * - 大きい画像は縦横400pxに制限
 * - 画像の縦横比率は元の画像と同じに維持
 */
public class ImageListExample {
    public static void main(String[] args) throws IOException {
        // 新規ワークブック作成
        try (Workbook workbook = new XSSFWorkbook()) {
            // シート作成
            Sheet sheet = workbook.createSheet("画像配置サンプル");

            // 画像ファイルのパス
            String[] imagePaths = {
                    "src/main/resources/test1.png",
                    "src/main/resources/test2.png",
                    "src/main/resources/test3.png",
                    "src/main/resources/test4.png",
                    "src/main/resources/test5.png"
            };

            // 画像を挿入
            insertImagesVertically(workbook, sheet, imagePaths);

            // ファイルに保存
            new File("output").mkdirs(); // outputディレクトリを作成
            String outputFilePath = "output/image_list.xlsx";
            try (FileOutputStream fileOut = new FileOutputStream(outputFilePath)) {
                workbook.write(fileOut);
                System.out.println("画像配置サンプルExcelファイルを作成しました。ファイルパス: " + outputFilePath);
            }
        }
    }

    /**
     * 複数の画像を縦に並べて挿入します。
     * 
     * @param workbook   ワークブック
     * @param sheet      シート
     * @param imagePaths 画像ファイルのパス配列
     * @throws IOException IO例外
     */
    private static void insertImagesVertically(Workbook workbook, Sheet sheet,
            String[] imagePaths) throws IOException {

        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();

        // 最大画像サイズ（ピクセル単位）
        int maxImageSize = 400; // 縦横最大400pxに制限

        // 各画像を挿入
        int currentRow = 1; // 開始行位置
        int imageSpacing = 2; // 画像間の間隔（行数）

        for (int i = 0; i < imagePaths.length; i++) {
            // 画像ファイルを読み込む
            File imageFile = new File(imagePaths[i]);
            BufferedImage originalImage = ImageIO.read(imageFile);
            int originalWidth = originalImage.getWidth();
            int originalHeight = originalImage.getHeight();

            // 画像の説明を追加
            Row descRow = sheet.createRow(currentRow);
            Cell descCell = descRow.createCell(0);
            descCell.setCellValue("画像 " + (i + 1) + " (元サイズ: " + originalWidth + "x" + originalHeight + "px)");
            currentRow += 1;

            // 最大サイズに合わせてスケーリング（ピクセル単位）
            double scale = 1.0;
            // 幅か高さが最大サイズを超える場合、縦横比を維持したままリサイズ
            if (originalWidth > maxImageSize || originalHeight > maxImageSize) {
                double scaleWidth = maxImageSize / (double) originalWidth;
                double scaleHeight = maxImageSize / (double) originalHeight;
                scale = Math.min(scaleWidth, scaleHeight);
            }

            int scaledWidth = (int) (originalWidth * scale);
            int scaledHeight = (int) (originalHeight * scale);

            // 実際に画像をリサイズ
            BufferedImage resizedImage = originalImage;
            if (scale < 1.0) {
                resizedImage = new BufferedImage(scaledWidth, scaledHeight, BufferedImage.TYPE_INT_ARGB);
                Graphics2D g = resizedImage.createGraphics();
                g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                g.drawImage(originalImage, 0, 0, scaledWidth, scaledHeight, null);
                g.dispose();
            }

            // リサイズした画像をバイト配列に変換
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ImageIO.write(resizedImage, "png", baos);
            byte[] imageBytes = baos.toByteArray();

            // バイト配列からExcelに画像を追加
            int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);

            // 画像の行数を計算（1行あたり20pxで概算）
            int rowSpan = Math.max(1, (int) Math.ceil(scaledHeight / 20.0));

            // 画像の開始行を記録
            int imageStartRow = currentRow;

            // セル幅を画像サイズに合わせて調整（1列あたり100pxで概算）
            int colSpan = Math.max(1, (int) Math.ceil(scaledWidth / 100.0));

            // 画像の位置とサイズを設定
            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setCol1(1); // B列に配置
            anchor.setRow1(imageStartRow);
            anchor.setCol2(1 + colSpan);
            anchor.setRow2(imageStartRow + rowSpan);

            // アンカーの種類を設定 - サイズを維持
            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

            // オフセットをクリア
            anchor.setDx1(0);
            anchor.setDy1(0);
            anchor.setDx2(0);
            anchor.setDy2(0);

            // 画像を描画
            Picture picture = drawing.createPicture(anchor, pictureIdx);
            picture.resize();

            // 次の画像のための位置調整
            currentRow = imageStartRow + rowSpan + imageSpacing;
        }
    }
}