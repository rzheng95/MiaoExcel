package com.rzheng.excel.util;


import java.io.File;
import java.util.Arrays;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.http.HttpEntity;
import org.apache.http.client.config.CookieSpecs;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.entity.mime.content.FileBody;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;

// Only 47 conversions left, limited to 50 for free account.
public class PDFToExcel_PDFTables {
	private static List<String> formats = Arrays.asList(new String[] { "csv", "xml", "xlsx-single", "xlsx-multiple" });

	public static void main(String[] args) throws Exception {

		final String apiKey = "e8rzszut6a4h";
		final String format = "xlsx-single";
		final String pdfFilename = "PI.pdf";

		if (!formats.contains(format)) {
			System.out.println("Invalid output format: \"" + format + "\"");
			System.exit(1);
		}

		// Avoid cookie warning with default cookie configuration
		RequestConfig globalConfig = RequestConfig.custom().setCookieSpec(CookieSpecs.STANDARD).build();

		File inputFile = new File(pdfFilename);

		if (!inputFile.canRead()) {
			System.out.println("Can't read input PDF file: \"" + pdfFilename + "\"");
			System.exit(1);
		}

		try (CloseableHttpClient httpclient = HttpClients.custom().setDefaultRequestConfig(globalConfig).build()) {
			HttpPost httppost = new HttpPost("https://pdftables.com/api?format=" + format + "&key=" + apiKey);
			FileBody fileBody = new FileBody(inputFile);

			HttpEntity requestBody = MultipartEntityBuilder.create().addPart("f", fileBody).build();
			httppost.setEntity(requestBody);

			System.out.println("Sending request");

			try (CloseableHttpResponse response = httpclient.execute(httppost)) {
				if (response.getStatusLine().getStatusCode() != 200) {
					System.out.println(response.getStatusLine());
					System.exit(1);
				}
				HttpEntity resEntity = response.getEntity();
				if (resEntity != null) {
					final String outputFilename = getOutputFilename(pdfFilename, format.replaceFirst("-.*$", ""));
					System.out.println("Writing output to " + outputFilename);

					final File outputFile = new File(outputFilename);
					FileUtils.copyToFile(resEntity.getContent(), outputFile);
				} else {
					System.out.println("Error: file missing from response");
					System.exit(1);
				}
			}
		}
	}

	private static String getOutputFilename(String pdfFilename, String suffix) {
		if (pdfFilename.length() >= 5 && pdfFilename.toLowerCase().endsWith(".pdf")) {
			return pdfFilename.substring(0, pdfFilename.length() - 4) + "." + suffix;
		} else {
			return pdfFilename + "." + suffix;
		}
	}
}
