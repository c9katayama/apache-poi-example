package poi_example.powerpoint;

import java.awt.Rectangle;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

/**
 * PowerPointファイルに画像を挿入するサンプル。
 */
public class ImageExample {
    public static void main(String[] args) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            // スライドの作成
            XSLFSlide slide = ppt.createSlide();

            // タイトルの追加
            XSLFTextBox title = slide.createTextBox();
            title.setAnchor(new Rectangle(50, 50, 500, 50));
            XSLFTextParagraph p = title.addNewTextParagraph();
            XSLFTextRun r = p.addNewTextRun();
            r.setText("画像を含むスライド");
            r.setFontSize(24.0);

            // 画像の挿入
            try (InputStream imageStream = ImageExample.class.getResourceAsStream("/sample_image.png")) {
                byte[] pictureData = imageStream.readAllBytes();

                XSLFPictureData xslfPictureData = ppt.addPicture(
                        pictureData,
                        XSLFPictureData.PictureType.PNG);
                XSLFPictureShape picture = slide.createPicture(xslfPictureData);
                picture.setAnchor(new Rectangle(100, 100, 400, 400));
            }

            // ファイルに保存
            new File("output").mkdirs();
            String outputFilePath = "output/image_slide_example.pptx";
            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                ppt.write(out);
                System.out.println("画像を含むPowerPointファイルを作成しました。ファイルパス: " + outputFilePath);
            }
        }
    }
}