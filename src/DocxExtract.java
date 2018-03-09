import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.HashSet;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.json.simple.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;

/**
 * Walk thru a folder of docx files, extract specified sections, text and context. Generate json and text file for each extracted unit.
 * Created Json and text files to be used for WDS and WKS.
 * 
 * @author muthukumaran
 *
 */
public class DocxExtract {

	private static int LVL_SPLIT = 2;
	
	//should be >=  LVL_SPLIT
	private static int LVL_CTX = 3;
	
	//should be >=  LVL_SPLIT
	private static HashSet<Integer> CONTENT_LVLS = new HashSet<Integer>();
	static{
		CONTENT_LVLS.add(Integer.valueOf(LVL_SPLIT));
		CONTENT_LVLS.add(Integer.valueOf(LVL_CTX));
	}

	private static String SRC_FOLDER = "/Users/muthukumaran/Documents/Projects/cg";
	private static String OUT_FOLDER = "/Users/muthukumaran/Documents/Projects/cg_docs";
	private static Logger logger = Logger.getLogger("Test Case Extractor");

	int extracted_tc = 0;
	int omitted_tc = 0;


	public static void main(String[] args) throws Exception {
			//new Test().getDocxByNumLvl("/Users/muthukumaran/Downloads/Testdoc.docx");
		DocxExtract t = new DocxExtract();
		Files.walk(Paths.get(SRC_FOLDER))
	     .filter(p -> p.toString().endsWith(".docx"))
	     .map(p -> p.toString())
	     .distinct()
	     .forEach(p -> {
			try {
				t.getDocxByNumLvl(p);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		});
			
		logger.log(Level.INFO, "Parsing Complete, Extracted "+t.extracted_tc+ ". Omitted "+t.omitted_tc );

	}
	
	/**
	 * @param inputFileName
	 * @param outputFileName
	 * @throws Exception
	 */
	public void getDocxByNumLvl(String inputFileName) throws Exception {

		String runningText = "";
		String ctx = "";
		String title = "";

		XWPFDocument docx = new XWPFDocument(new FileInputStream(new File(inputFileName)));
		XWPFStyles styles = docx.getStyles();

		boolean capturingText = false;
		boolean isContext = false;

		logger.log(Level.INFO, "Parsing document : "+inputFileName);

		for (XWPFParagraph para : docx.getParagraphs()) {

			int pLvl_val = -1;
			String paraText = "";
			
			//get paragraph text.
			for (XWPFRun run : para.getRuns()) {
				paraText += run.text();
			}

			//if paragraph is a heading
			if (null != para.getStyle()) {

				CTDecimalNumber oLvl = null;
				try {
					oLvl = styles.getStyle(para.getStyle()).getCTStyle().getPPr().getOutlineLvl();
				} catch (Exception x) {
					//ignore and continue
					continue;
				}

				if (null != oLvl) {
					pLvl_val = oLvl.getVal().intValue();

					if (pLvl_val < LVL_SPLIT) {
						capturingText = false;
						continue;
					}

					if (pLvl_val == LVL_SPLIT) {
						capturingText = true;
						
						if (!runningText.isEmpty()) {
							createJson(runningText, ctx, title);
						}
						title=paraText;
						runningText = "";
						ctx="";
					}

					if (pLvl_val == LVL_CTX) {
						isContext = true;
					}else{
						isContext = false;
					}
				}
			}

			if (capturingText) {
				runningText += " "+paraText;
				if (pLvl_val < 0 && isContext && ctx.isEmpty()) {
					ctx = paraText;
				}
			}
		}

		createJson(runningText, ctx, title);
		docx.close();
	}

	/**
	 * @param e_txt
	 * @param e_ctx
	 * @param e_title
	 */
	private void createJson(String e_txt, String e_ctx, String e_title) {
		JSONObject json = new JSONObject();
		json.put("procedure", e_txt);
		json.put("ctx", e_ctx);
		json.put("title", e_title);
		
		if (e_txt.isEmpty() || e_ctx.isEmpty() || e_title.isEmpty()){
			logger.log(Level.INFO, "UnExpected Blank "+json.toJSONString());
			omitted_tc++;
		}else{
			try {
				long ms = System.currentTimeMillis();
				Files.write(Paths.get(OUT_FOLDER + "//tc_" + ms + ".json"), json.toJSONString().getBytes());
				Files.write(Paths.get(OUT_FOLDER + "//tc_" + ms + ".txt"), e_txt.getBytes(Charset.forName("UTF-8")), StandardOpenOption.CREATE);
				extracted_tc++;
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
}