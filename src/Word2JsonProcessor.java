import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;

public class Word2JsonProcessor {

	JSONArray testDocs = new JSONArray();

	private static int SPLIT_OUTLINE_LVL = 2;
	private static int SPLITDOC_CONTEXT_LVL = 3;
	
	private static int fontSizeBarrier = 11;

	public static void main(String[] args) throws Exception {

		String inputFileName = "/Users/muthukumaran/Downloads/Testdoc.docx";
		String outputFileName = "documents/TestProcedures.json";

		if (inputFileName.endsWith("docx")) {
			new Word2JsonProcessor().getDocxByNumLvl(inputFileName, outputFileName);
		} else if (inputFileName.endsWith("doc")) {
			new Word2JsonProcessor().performDocConv(inputFileName, outputFileName);
		} else {
			System.err.println("This input file typr is not handled. ");
		}
	}

	public void performDocConv(String inputFileName, String outputFileName)  throws Exception{

		HWPFDocument doc = new  HWPFDocument(new FileInputStream(inputFileName));
		int runningFontSize = -1;
		String runningText = "";
		JSONArray testDocs = new JSONArray();
		JSONObject context = new JSONObject();

		final Range range = doc.getRange();
		for (int k = 0; k < range.numParagraphs(); k++) {
			final org.apache.poi.hwpf.usermodel.Paragraph paragraph = range.getParagraph(k);
			for (int j = 0; j < paragraph.numCharacterRuns(); j++) {
				final org.apache.poi.hwpf.usermodel.CharacterRun cr = paragraph.getCharacterRun(j);
				
				int fontSize = cr.getFontSize()/2;
				
				if (runningFontSize <0) {
					runningFontSize = fontSize;
					runningText = cr.text();
				}
				else {
					if (runningFontSize != fontSize) {
						if (runningFontSize <= fontSizeBarrier) {
							JSONObject testBlock = new JSONObject();
							testBlock.put("testBlock", runningText);
							testBlock.put("context", new JSONObject( context));
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
		doc.close();
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
			for (XWPFRun run:para.getRuns()) {
				String paraStyle = para.getStyle();
				int fontSize = run.getFontSize();
				if (fontSize == -1) {
					if( paraStyle != null){
						fontSize = (docx.getStyles().getStyle(paraStyle).getCTStyle().getRPr().getSz().getVal().intValue())/2;
					}
					else {
						fontSize = docx.getStyles().getDefaultRunStyle().getFontSize();
					}
				}				
				if (runningFontSize <0) {
					runningFontSize = fontSize;
					runningText = run.text();
				}
				else {
					if (runningFontSize != fontSize) {
						if (runningFontSize <= fontSizeBarrier) {
							JSONObject testBlock = new JSONObject();
							testBlock.put("testBlock", runningText);
							testBlock.put("context", new JSONObject( context));
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
	/**
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	public void getDocxByNumLvl(String inputFileName, String outputFileName) throws Exception {

		Map<String,String> ctxMap = new HashMap<String,String>();
		String runningText = "";

		XWPFDocument docx = new XWPFDocument(new FileInputStream(inputFileName));
		XWPFStyles styles = docx.getStyles();

		boolean append_paragraphContent = false;
		boolean addParagraphContentAsContext = false;

		for (XWPFParagraph para : docx.getParagraphs()) {

			int outlineLvl = -1;
			String paraStyle = para.getStyle();

			String cText = "";
			for (XWPFRun run : para.getRuns()) {
				cText += run.text()+" ";
			}

			if (null != paraStyle) {

				CTDecimalNumber oLvl = styles.getStyle(paraStyle).getCTStyle().getPPr().getOutlineLvl();

				if (null != oLvl) {
					outlineLvl = oLvl.getVal().intValue();

					if (outlineLvl < SPLIT_OUTLINE_LVL) {
						append_paragraphContent = false;
						continue;
					}
					
					if (outlineLvl == SPLIT_OUTLINE_LVL) {
						append_paragraphContent = true;
						if (!runningText.isEmpty())
							populateJson(runningText,ctxMap);
							runningText = "";
							ctxMap.put("title", cText);
					}

					if (outlineLvl == SPLITDOC_CONTEXT_LVL) {
						addParagraphContentAsContext = true;
					} else {
						addParagraphContentAsContext = false;
					}
				}
			}

			if (append_paragraphContent) {
				runningText += cText;
				if (outlineLvl < 0 && addParagraphContentAsContext) {
					ctxMap.put("ctx", cText);
				}
			}
		}

		populateJson(runningText,ctxMap);

		docx.close();
		System.out.println(testDocs.toJSONString());
		System.setOut(new PrintStream(new FileOutputStream(new File(outputFileName))));
		System.out.println(testDocs.toJSONString());
	}

	private void populateJson(String text, Map<String, String> ctxMap) {
		if (!text.isEmpty()) {
			JSONObject tc = new JSONObject();
			tc.put("testcase", text);

			for (Map.Entry<String, String> entry : ctxMap.entrySet()) {
			    String key = entry.getKey();
			    String value = entry.getValue();
			    tc.put(key, value);
			}
			testDocs.add(tc);
		}
	}
}
