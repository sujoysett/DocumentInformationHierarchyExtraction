import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.Iterator;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.json.simple.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;

/**
 * Walk thru a folder of docx files, extract specified sections, text and
 * context. Generate json and text file for each extracted unit. Created Json
 * and text files to be used for WDS and WKS.
 * 
 * @author muthukumaran
 *
 */
public class DocxExtract {

	private static Properties config = new Properties();
	static {
		try {
			config.load(Files.newInputStream(Paths.get("config/config.properties")));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	private static Logger logger = Logger.getLogger("Test Case Extractor");

	int extracted_tc = 0;
	int omitted_tc = 0;
	int LVL_SPLIT = Integer.parseInt(config.getProperty("LVL_SPLIT"));
	int LVL_CTX = Integer.parseInt(config.getProperty("LVL_CTX"));
	String SRC_FOLDER = config.getProperty("SRC_FOLDER");
	String OUT_FOLDER = config.getProperty("OUT_FOLDER");

	public DocxExtract() {
		try {
			Files.walk(Paths.get(SRC_FOLDER)).filter(p -> p.toString().endsWith(".docx")).map(p -> p.toString()).distinct()
					.forEach(p -> {
						try {
							this.extractSectionContent(p, LVL_SPLIT);
						} catch (Exception e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					});
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		logger.log(Level.INFO, "Parsing Complete, Extracted " + this.extracted_tc + ". Omitted " + this.omitted_tc);

	}

	public static void main(String[] args) throws Exception {
		// new
		// Test().getDocxByNumLvl("/Users/muthukumaran/Downloads/Testdoc.docx");
		DocxExtract t = new DocxExtract();
	}

	/**
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	public void extractSectionContent(String inputFileName, int S_LVL) throws Exception {

		String runningText = "";
		String ctx = "";
		String title = "";

		XWPFDocument docx = new XWPFDocument(new FileInputStream(new File(inputFileName)));
		XWPFStyles styles = docx.getStyles();

		boolean capturingText = false;
		boolean capturingTables = false;

		boolean isContext = false;

		logger.log(Level.INFO, "Parsing document : " + inputFileName);

		Iterator<IBodyElement> iter = docx.getBodyElementsIterator();
		while (iter.hasNext()) {
			IBodyElement elem = iter.next();
			if (elem instanceof XWPFParagraph) {

				XWPFParagraph para = (XWPFParagraph) elem;

				int pLvl_val = -1;
				String paraText = "";

				// get paragraph text.
				for (XWPFRun run : para.getRuns()) {
					paraText += run.text();
				}

				// if paragraph is a heading
				if (null != para.getStyle()) {

					CTDecimalNumber oLvl = null;
					try {
						oLvl = styles.getStyle(para.getStyle()).getCTStyle().getPPr().getOutlineLvl();
					} catch (Exception x) {
						// ignore and continue
						continue;
					}

					if (null != oLvl) {
						pLvl_val = oLvl.getVal().intValue();

						if (pLvl_val < S_LVL) {
							capturingText = false;
							capturingTables = false;
							continue;
						}

						if (pLvl_val == S_LVL) {
							capturingText = true;
							capturingTables = true;

							if (!runningText.isEmpty()) {
								createFiles(runningText, ctx, title);
							}
							title = paraText;
							runningText = "";
							ctx = "";
						}

						if (pLvl_val == LVL_CTX) {
							isContext = true;
						} else {
							isContext = false;
						}
					}
				}

				if (capturingText) {
					runningText += System.lineSeparator() + paraText;
					if (pLvl_val < 0 && isContext && ctx.isEmpty()) {
						ctx = paraText;
					}
				}
			} else if (capturingTables && elem instanceof XWPFTable) {

				runningText += System.lineSeparator();

				XWPFTable table = (XWPFTable) elem;
				for (XWPFTableRow row : table.getRows()) {
					runningText += System.lineSeparator();
					for (XWPFTableCell cell : row.getTableCells()) {
						runningText += cell.getText() + "\t";
					}
				}
			}
		}

		createFiles(runningText, ctx, title);
		docx.close();
	}

	/**
	 * @param e_txt
	 * @param e_ctx
	 * @param e_title
	 */
	private void createFiles(String e_txt, String e_ctx, String e_title) {
		JSONObject json = new JSONObject();
		json.put("procedure", e_txt);
		json.put("ctx", e_ctx);
		json.put("title", e_title);

		if (e_txt.isEmpty() || e_ctx.isEmpty() || e_title.isEmpty()) {
			logger.log(Level.INFO, "UnExpected Blank " + json.toJSONString());
			omitted_tc++;
		} else {
			try {
				String ms = "" + extracted_tc;
				Files.write(Paths.get(OUT_FOLDER + "//tc_" + ms + ".json"), json.toJSONString().getBytes());
				Files.write(Paths.get(OUT_FOLDER + "//tc_" + ms + ".txt"), e_txt.getBytes(Charset.forName("UTF-8")),
						StandardOpenOption.CREATE);
				extracted_tc++;
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
}