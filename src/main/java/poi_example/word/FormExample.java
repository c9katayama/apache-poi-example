package poi_example.word;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * Docxで用意したフォームに、Javaのモデルデータを差し込むサンプル。
 * フォームのテンプレートはform-template.docxを使用する。
 * 
 */
public class FormExample {

	// form-template.docxを読み込んで、フォームを作成する
	public static class FormModel {
		public String num;
		public String year;
		public String month;
		public String day;
		public String address;
		public String companyName;
		public String name;
		public String description;
		public String reason;
	}

	public static void main(String[] args) throws IOException {
		InputStream fis = FormExample.class.getResourceAsStream("/form-template.docx");

		try (XWPFDocument document = new XWPFDocument(fis)) {
			FormModel formModel = new FormModel();
			formModel.num = "123";
			formModel.year = "2025";
			formModel.month = "1";
			formModel.day = "1";
			formModel.address = "東京都千代田区永田町1-2-3";
			formModel.companyName = "株式会社テスト";
			formModel.name = "山田太郎";
			formModel.description = "変更内容の詳細について";
			formModel.reason = "変更理由の詳細について";
			byte[] filledForm = createForm(document, formModel);

			// ファイルに保存
			new File("output").mkdirs();
			try (FileOutputStream out = new FileOutputStream("output/form.docx")) {
				out.write(filledForm);
			}
		}
		System.out.println("Formファイルを作成しました。");
	}

	// フォームを作成する
	// フォーム内の${}文字列をModelのプロパティ値に置換する
	private static byte[] createForm(XWPFDocument document, Object formModel) throws IOException {
		Map<String, String> replacements = toPropertyMapMap(formModel);
		for (XWPFParagraph p : document.getParagraphs()) {
			replaceInParagraph(p, replacements);
		}
		for (XWPFTable tbl : document.getTables()) {
			for (XWPFTableRow row : tbl.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					for (XWPFParagraph p : cell.getParagraphs()) {
						replaceInParagraph(p, replacements);
					}
				}
			}
		}
		try (ByteArrayOutputStream bout = new ByteArrayOutputStream()) {
			document.write(bout);
			document.close();
			return bout.toByteArray();
		}
	}

	// オブジェクトのフィールドをMapに変換する。キーは${name}の形式になるため、フォームの置換に使用する。
	private static Map<String, String> toPropertyMapMap(Object o) {
		Map<String, String> map = new HashMap<>();
		Field[] fields = o.getClass().getFields();
		for (Field f : fields) {
			if (Modifier.isStatic(f.getModifiers())) {
				continue;
			}
			String name = f.getName();
			Object value;
			try {
				value = f.get(o);
			} catch (IllegalArgumentException | IllegalAccessException e) {
				throw new RuntimeException(e);
			}
			if (value != null) {
				if (value instanceof List) {
					List<?> list = (List<?>) value;
					for (int i = 0; i < list.size(); i++) {
						map.put("${" + name + i + "}", list.get(i).toString());
					}
				} else if (value instanceof String[]) {
					String[] list = (String[]) value;
					for (int i = 0; i < list.length; i++) {
						map.put("${" + name + i + "}", (list[i] == null ? "" : list[i].toString()));
					}
				} else if (value instanceof String) {
					map.put("${" + name + "}", value.toString());
				} else {
					throw new RuntimeException(
							"Unsupported type: " + value.getClass().getName() + " field:" + f.getName());
				}
			}
		}
		return map;
	}

	// Word doc中の${}文字列を置換する
	private static void replaceInParagraph(XWPFParagraph p, Map<String, String> replacements) {
		// 複数のフラグメントに分かれてる場合があるので、連結する
		StringBuffer buf = new StringBuffer();
		for (XWPFRun r : p.getRuns()) {
			String text = r.getText(0);
			if (text != null) {
				buf.append(text);
			}
		}
		String text = buf.toString();
		boolean replace = false;
		for (Map.Entry<String, String> replEntry : replacements.entrySet()) {
			if (text.contains(replEntry.getKey())) {
				text = text.replace(replEntry.getKey(), replEntry.getValue());
				replace = true;
			}
		}
		// 置き換えた場合、最初のフラグメントに連結/置換文字列をいれて、残りは空にする
		if (replace) {
			List<XWPFRun> runs = p.getRuns();
			for (int i = 0; i < runs.size(); i++) {
				XWPFRun run = runs.get(i);
				if (i == 0) {
					run.setText(text, 0);
				} else {
					run.setText("", 0);
				}
			}
		}
	}
}