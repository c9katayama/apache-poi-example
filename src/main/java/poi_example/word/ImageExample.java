package poi_example.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Wordドキュメントに画像を挿入するサンプル。
 */
public class ImageExample {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        try (XWPFDocument document = new XWPFDocument()) {
            // 段落の作成
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);

            // 画像の挿入
            try (FileInputStream imageStream = new FileInputStream("sample_image.png")) {
                XWPFRun run = paragraph.createRun();
                run.addPicture(
                        imageStream,
                        Document.PICTURE_TYPE_PNG,
                        "sample_image.png",
                        Units.toEMU(300), // 幅
                        Units.toEMU(200) // 高さ
                );
            }

            // 画像の説明文を追加
            XWPFParagraph caption = document.createParagraph();
            caption.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun captionRun = caption.createRun();
            captionRun.setText("サンプル画像");
            captionRun.setFontSize(10);
            captionRun.setItalic(true);

            // ファイルに保存
            new File("output").mkdirs();
            String outputFilePath = "output/image_example.docx";
            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                document.write(out);
                System.out.println("画像を含むWordファイルを作成しました。ファイルパス: " + outputFilePath);
            }
        }
    }
}