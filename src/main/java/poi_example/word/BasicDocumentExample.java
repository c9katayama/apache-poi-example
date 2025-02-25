package poi_example.word;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Wordドキュメントを作成し、段落を入れて保存するサンプル。
 */
public class BasicDocumentExample {
	public static void main(String[] args) throws IOException {
		try (XWPFDocument document = new XWPFDocument()) {
			// 段落の作成
			XWPFParagraph paragraph = document.createParagraph();
			XWPFRun run = paragraph.createRun();
			run.setText("Hello Apache POI Word!");
			run.setBold(true);
			run.setFontSize(14);

			// 新しい段落を作成
			paragraph = document.createParagraph();
			run = paragraph.createRun();
			run.setText("This is a new paragraph.");

			// ファイルに保存
			new File("output").mkdirs();
			String outputFilePath = "output/basic_document.docx";
			try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
				document.write(out);
				System.out.println("Wordファイルを作成しました。ファイルパス: " + outputFilePath);
			}
		}
	}
}