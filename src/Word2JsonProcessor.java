import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.PrintStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

public class Word2JsonProcessor {

	private static int fontSizeBarrier = 11;

	public static void main(String[] args) throws Exception {

		// String inputFileName = "/Users/muthukumaran/Downloads/Testdoc.docx";
		String inputFileName = "C://Users//jagvenug//Desktop//Watson-POT//CSCS V2 DVP Real Time Data.docx";

		String outputFileName = "documents/TestProcedures.json";

		if (inputFileName.endsWith("docx")) {
			new Word2JsonProcessor().performDocxConv(inputFileName, outputFileName);
		} else if (inputFileName.endsWith("doc")) {
			new Word2JsonProcessor().performDocConv(inputFileName, outputFileName);
		} else {
			System.err.println("This input file typr is not handled. ");
		}
	}

	public void performDocConv(String inputFileName, String outputFileName) throws Exception {

		HWPFDocument doc = new HWPFDocument(new FileInputStream(inputFileName));
		int runningFontSize = -1;
		String runningText = "";
		JSONArray testDocs = new JSONArray();
		JSONObject context = new JSONObject();

		final Range range = doc.getRange();
		for (int k = 0; k < range.numParagraphs(); k++) {
			final org.apache.poi.hwpf.usermodel.Paragraph paragraph = range.getParagraph(k);
			for (int j = 0; j < paragraph.numCharacterRuns(); j++) {
				final org.apache.poi.hwpf.usermodel.CharacterRun cr = paragraph.getCharacterRun(j);

				int fontSize = cr.getFontSize() / 2;

				if (runningFontSize < 0) {
					runningFontSize = fontSize;
					runningText = cr.text();
				} else {
					if (runningFontSize != fontSize) {
						if (runningFontSize <= fontSizeBarrier) {
							JSONObject testBlock = new JSONObject();
							testBlock.put("testBlock", runningText);
							testBlock.put("context", new JSONObject(context));
							testDocs.add(testBlock);
						} else if (runningFontSize > fontSizeBarrier) {
							context.put(runningFontSize, runningText);
						}
						runningText = "";
						runningFontSize = fontSize;
					}
					runningText = runningText + cr.text();
				}
			}
		}
		//doc.close();
		System.setOut(new PrintStream(new FileOutputStream(new File(outputFileName))));
		System.out.println(testDocs.toJSONString());
	}

	public void performDocxConv(String inputFileName, String outputFileName) throws Exception {

		XWPFDocument docx = new XWPFDocument(new FileInputStream(inputFileName));
		int runningFontSize = -1;
		String runningText = "";
		JSONArray testDocs = new JSONArray();
		JSONObject context = new JSONObject();

		for (XWPFParagraph para : docx.getParagraphs()) {
			for (XWPFRun run : para.getRuns()) {
				String paraStyle = para.getStyle();
				int fontSize = run.getFontSize();
				if (fontSize == -1) {
					if (paraStyle != null) {
						fontSize = (docx.getStyles().getStyle(paraStyle).getCTStyle().getRPr().getSz().getVal()
								.intValue()) / 2;
					} else {
						fontSize = docx.getStyles().getDefaultRunStyle().getFontSize();
					}
				}
				if (runningFontSize < 0) {
					runningFontSize = fontSize;
					runningText = run.text();
				} else {
					if (runningFontSize != fontSize) {
						if (runningFontSize <= fontSizeBarrier) {
							JSONObject testBlock = new JSONObject();
							testBlock.put("testBlock", runningText);
							testBlock.put("context", new JSONObject(context));
							testDocs.add(testBlock);
						} else if (runningFontSize > fontSizeBarrier) {
							context.put(runningFontSize, runningText);
						}
						runningText = "";
						runningFontSize = fontSize;
					}
					runningText = runningText + run.text();
				}
			}
		}
		docx.close();
		System.setOut(new PrintStream(new FileOutputStream(new File(outputFileName))));
		System.out.println(testDocs.toJSONString());
	}
}
