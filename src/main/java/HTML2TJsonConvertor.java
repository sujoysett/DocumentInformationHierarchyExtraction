package main.java;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.regex.Pattern;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

public class HTML2TJsonConvertor {
	private static int fontSizeBarrier = 13;

	private static Pattern procedurePattern = Pattern.compile("(complete the following)");

	private static JSONArray productDocs = new JSONArray();

	public static void main(String []s) throws Exception {
		Document htmlDocument = Jsoup.parse(new FileInputStream(new File("documents/full_doc_output20180227.html")),"utf-8", ""); 
		Elements spans = htmlDocument.select("span");

		boolean productDescriptionStarted = false;
		String productDescription = "";

		JSONObject context = new JSONObject();
		for (int spanIndex = 0; spanIndex < spans.size(); spanIndex++) {
			String style = spans.get(spanIndex).attr("style");
			if (style.contains("font-size")) {
				String fontSizeStr = spans.get(spanIndex).attr("style").split("(font-size:)")[1];
				int fontSizeNum = Integer.parseInt(fontSizeStr.split("px")[0]);
				String text = spans.get(spanIndex).text();

				if (fontSizeNum > fontSizeBarrier) {
					context.put(fontSizeStr, text);
					if (productDescriptionStarted) {
						productDescriptionStarted = false;

						JSONObject productDoc = new JSONObject();
						productDoc.put("procedure", productDescription);
						productDoc.put("context", new JSONObject( context));
						productDocs.add(productDoc);
						productDescription = "";
						
						System.out.println("Procedure Count = " + productDocs.size());
					}
				}

				if (fontSizeNum <= fontSizeBarrier) {
					if(procedurePattern.matcher(text.toLowerCase()).find()) {
						productDescriptionStarted = true;
						productDescription = "";
					}
					if(productDescriptionStarted) {
						productDescription =productDescription + " " + text;
					} 
				}
			}
			//System.out.println(productDescription);
		}
		System.setOut(new PrintStream(new FileOutputStream(new File("documents/ProductProcedures.json"))));
		System.out.println(productDocs.toJSONString());
	}
}
