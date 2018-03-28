import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.OutputStreamWriter;
import java.io.PrintStream;
import java.util.regex.Pattern;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

public class HTML2TJsonConvertor {
	private static int headerToTextFontSizeBarrier = 13;
	private static int unwantedTextFontSizeBarrier = 10;
	private static String htmlDoc = "documents/full_doc_output20180308.html";

	private static Pattern procedurePattern = Pattern.compile("(complete the following)");
	private static JSONArray productDocs = new JSONArray();
	
	private static String textOutputFolder = "ProductProcedureText";
	private static String jsonOutputFolder = "ProductProcedureJson";

	public static void main(String []s) {
		try {
			analyzeHtml();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void analyzeHtml() throws Exception {
		Document htmlDocument = Jsoup.parse(new FileInputStream(new File(htmlDoc)),"utf-8", ""); 
		Elements spans = htmlDocument.select("span");

		boolean productDescriptionStarted = false;
		String productDescription = "";
		String immediatePrecedingHigherFontString = "";
		int immediatePrecedingFontSize = 0;


		JSONObject context = new JSONObject();
		for (int spanIndex = 0; spanIndex < spans.size(); spanIndex++) {
			String style = spans.get(spanIndex).attr("style");
			if (style.contains("font-size")) {
				String fontSizeStr = spans.get(spanIndex).attr("style").split("(font-size:)")[1];
				int fontSizeNum = Integer.parseInt(fontSizeStr.split("px")[0]);
				String text = spans.get(spanIndex).text();

				if (fontSizeNum <= unwantedTextFontSizeBarrier) {
					// reject such text, these are usually headers and footers.
					continue;
				}

				if (fontSizeNum > headerToTextFontSizeBarrier) {

					if (productDescriptionStarted) {
						productDescriptionStarted = false;

						String productDescriptionBody = "The task name is \"" + immediatePrecedingHigherFontString +"\""
								+ System.lineSeparator() + System.lineSeparator() +productDescription;


						JSONObject productDoc = new JSONObject();
						productDoc.put("procedure", productDescriptionBody);
						productDoc.put("full_context", new JSONObject( context));
						productDoc.put("ctx", "The task name is \"" + immediatePrecedingHigherFontString + "\"");
						productDoc.put("title", immediatePrecedingHigherFontString);
						productDocs.add(productDoc);

						writeProcToIndividualTextFile(productDocs.size()+1, productDescriptionBody);
						writeProcToIndividualJsonFile(productDocs.size()+1, productDoc);
						productDescription = "";

						System.out.println("Procedure Count = " + productDocs.size());
					}

					if (fontSizeNum == immediatePrecedingFontSize) {
						immediatePrecedingHigherFontString = immediatePrecedingHigherFontString + " " + text;
					} else {
						immediatePrecedingHigherFontString = text;
					}
					context.put(fontSizeStr, immediatePrecedingHigherFontString);
				}

				if (fontSizeNum <= headerToTextFontSizeBarrier && fontSizeNum > unwantedTextFontSizeBarrier) {
					if(procedurePattern.matcher(text.toLowerCase()).find()) {
						productDescriptionStarted = true;
						productDescription = "";
					}
					if(productDescriptionStarted) {
						productDescription = productDescription + " " + text;
					} 
				}
				immediatePrecedingFontSize = fontSizeNum;
			}
		}
		System.setOut(new PrintStream(new FileOutputStream(new File("documents/ProductProcedures.json"))));
		System.out.println(productDocs.toJSONString());
	}


	private static void writeProcToIndividualTextFile(int i, String productDescription) {
		BufferedWriter writer = null;
		try {
			FileOutputStream fileStream = new FileOutputStream(textOutputFolder + "/ProductUserManualStep"+ i+".txt");
			OutputStreamWriter writerStream = new OutputStreamWriter(fileStream,"UTF-8");	
			writer = new BufferedWriter(writerStream);
			writer.write(productDescription);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				writer.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
	
	private static void writeProcToIndividualJsonFile(int i, JSONObject productDoc) {
		BufferedWriter writer = null;
		try {
			FileOutputStream fileStream = new FileOutputStream(jsonOutputFolder + "/ProductUserManualStep"+ i+".json");
			OutputStreamWriter writerStream = new OutputStreamWriter(fileStream,"UTF-8");	
			writer = new BufferedWriter(writerStream);
			writer.write(productDoc.toJSONString());
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				writer.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
}
