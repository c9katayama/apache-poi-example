package poi_example.powerpoint;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

/**
 * PowerPointファイルを作成するサンプル。
 */
public class BasicPresentationExample {
    public static void main(String[] args) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            // スライドの作成
            XSLFSlide slide = ppt.createSlide();

            // タイトルの追加
            XSLFTextBox title = slide.createTextBox();
            title.setAnchor(new java.awt.Rectangle(50, 50, 400, 50));

            XSLFTextParagraph p = title.addNewTextParagraph();
            XSLFTextRun r = p.addNewTextRun();
            r.setText("Hello Apache POI PowerPoint!");
            r.setFontSize(32.0);

            // ファイルに保存
            new File("output").mkdirs();
            String outputFilePath = "output/presentation.pptx";
            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                ppt.write(out);
                System.out.println("PowerPointファイルを作成しました。ファイルパス: " + outputFilePath);
            }
        }
    }
}