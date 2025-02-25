package poi_example.powerpoint;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

/**
 * PowerPointファイルに図形を追加するサンプル。
 */
public class ShapeExample {
    public static void main(String[] args) throws IOException {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            // スライドの作成
            XSLFSlide slide = ppt.createSlide();

            // 四角形の追加
            XSLFAutoShape rectangle = slide.createAutoShape();
            rectangle.setShapeType(ShapeType.RECT);
            rectangle.setAnchor(new Rectangle(100, 100, 200, 100));
            rectangle.setFillColor(Color.BLUE);

            // 円の追加
            XSLFAutoShape circle = slide.createAutoShape();
            circle.setShapeType(ShapeType.ELLIPSE);
            circle.setAnchor(new Rectangle(350, 100, 100, 100));
            circle.setFillColor(Color.RED);

            // 矢印の追加
            XSLFAutoShape arrow = slide.createAutoShape();
            arrow.setShapeType(ShapeType.RIGHT_ARROW);
            arrow.setAnchor(new Rectangle(100, 250, 200, 50));
            arrow.setFillColor(Color.GREEN);

            // テキストボックスの追加
            XSLFTextBox textBox = slide.createTextBox();
            textBox.setAnchor(new Rectangle(350, 250, 200, 50));
            XSLFTextParagraph p = textBox.addNewTextParagraph();
            XSLFTextRun r = p.addNewTextRun();
            r.setText("図形のサンプル");
            r.setFontSize(14.0);

            // ファイルに保存
            new File("output").mkdirs();
            String outputFilePath = "output/shape_example.pptx";
            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                ppt.write(out);
                System.out.println("図形を含むPowerPointファイルを作成しました。ファイルパス: " + outputFilePath);
            }
        }
    }
}