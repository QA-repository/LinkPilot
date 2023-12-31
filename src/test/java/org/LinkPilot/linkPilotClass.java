package org.LinkPilot;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.net.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.*;

public class linkPilotClass {
    public void main(String[] args) {
        System.out.println("inside main");
        Link_Checker();
    }

    static long startTime;
    static long endTime;
    static String urlType = "internal";
    private static int scannedLinks = 0;

    @Test
    public void Link_Checker() {
        System.out.println("inside main");

        try {
            String startingUrl = System.getProperty("startingUrl", "defaultStartingUrl");
            String cookie = System.getProperty("cookie", "defaultCookieValue");

            Set<String> processedURLs = new HashSet<>();
            Workbook workbook = new XSSFWorkbook();
            Sheet resultsSheet = workbook.createSheet("Testing Result");

            crawlAndVerifyURL(startingUrl, cookie, resultsSheet, processedURLs);

            try (FileOutputStream fileOut = new FileOutputStream(System.getProperty("user.dir")+"/LatestResult")) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
                logError("Error writing workbook to file: " + e.getMessage());
                System.exit(-1);
            }

        } catch (Exception e) {
            e.printStackTrace();
            logError("Unhandled exception: " + e.getMessage());
            System.exit(-1);
        }
    }

    private static void crawlAndVerifyURL(String startingUrl, String cookie, @NotNull Sheet resultsSheet, Set<String> processedURLs) {
        Queue<String> queue = new LinkedList<>();
        Set<String> visitedURLs = new HashSet<>();

        queue.add(startingUrl);

        Row headerRow = resultsSheet.createRow(0);
        createCell(headerRow, 0, "URL");
        createCell(headerRow, 1, "URL Type");
        createCell(headerRow, 2, "linkPilotClass Status");
        createCell(headerRow, 3, "H1 tags number");
        createCell(headerRow, 4, "Response Time");
        createCell(headerRow, 5, "PDF Size");
        createCell(headerRow, 6, "Image Size");

        while (!queue.isEmpty()) {
            String currentUrl = queue.poll();

            if (processedURLs.contains(currentUrl)) {
                continue;
            }

            if (!isWellFormattedHttpUrl(currentUrl)) {
                System.out.println("URL: " + currentUrl + " is not well-formatted or not HTTP. Skipping...");
                continue;
            }

            boolean isSameDomain = isSameDomain(startingUrl, currentUrl);

            scannedLinks++;

            try {
                String[] apiResponse = apiResponse(currentUrl, cookie);
                String responseCode = apiResponse[0];
                String responseText = apiResponse[1];
                String result = getResult(currentUrl, responseCode, responseText, resultsSheet);

                Row row = resultsSheet.createRow(resultsSheet.getLastRowNum() + 1);
                Document document = Jsoup.connect(currentUrl).header("Cookie", cookie).get();
                Elements h1Tags = document.select("h1");
                int number_Of_H1Tags = h1Tags.size();
                String h1TagResult = String.valueOf(number_Of_H1Tags);
                createCell(row, 0, currentUrl);
                createCell(row, 1, urlType);
                createCell(row, 2, result);
                createCell(row, 3, h1TagResult);
                createCell(row, 4, String.valueOf(endTime - startTime));

                processedURLs.add(currentUrl);

                if (isSameDomain || !result.equals("Passed")) {
                    Elements links = document.select("a[href]");

                    for (Element link : links) {
                        String url = link.absUrl("href");

                        if (!isWellFormattedHttpUrl(url)) {
                            System.out.println("URL: " + url + " is not well-formatted or not HTTP. Skipping...");
                            continue;
                        }

                        if (!url.isEmpty() && visitedURLs.add(url)) {

                            if (isSameDomain(startingUrl, url)) {
                                queue.add(url);
                                urlType = "internal";
                            } else {
                                String[] externalApiResponse = apiResponse(url, cookie);
                                String externalResponseCode = externalApiResponse[0];
                                String externalResponseText = externalApiResponse[1];
                                String externalResult = getResult(url, externalResponseCode, externalResponseText, resultsSheet);

                                Row externalRow = resultsSheet.createRow(resultsSheet.getLastRowNum() + 1);
                                createCell(externalRow, 0, url);
                                createCell(externalRow, 1, "External");
                                createCell(externalRow, 2, externalResult);
                                createCell(externalRow, 3, h1TagResult);
                                createCell(externalRow, 4, String.valueOf(endTime - startTime));

                                if (externalResponseCode.equals("200") && externalResult.equals("Passed")) {
                                    verifyPDFAndImageSizes(url, externalRow, cookie);
                                }

                                processedURLs.add(url);

                                if (processedURLs.contains(url)) {
                                    urlType = "External";
                                }
                            }
                        }
                    }
                }
            } catch (IOException e) {
                logException(currentUrl, e);
                continue;
            }
        }
    }

    private static void verifyPDFAndImageSizes(String urlString, Row row, String cookie) {
        try {
            URL url = new URL(urlString);
            URLConnection connection = url.openConnection();
            connection.setRequestProperty("Cookie", cookie);

            if (connection instanceof HttpURLConnection) {
                HttpURLConnection httpURLConnection = (HttpURLConnection) connection;

                httpURLConnection.setRequestMethod("GET");

                try (InputStream inputStream = httpURLConnection.getInputStream()) {
                    String contentType = httpURLConnection.getContentType();

                    if (contentType != null && contentType.toLowerCase().contains("application/pdf")) {
                        verifyPDFResponse(urlString, httpURLConnection, row);
                    } else if (contentType != null && contentType.toLowerCase().startsWith("image/")) {
                        verifyImageResponse(urlString, httpURLConnection, row);
                    }
                } finally {
                    endTime = System.currentTimeMillis();
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void verifyPDFResponse(String urlString, HttpURLConnection httpURLConnection, Row row) {
        try (InputStream inputStream = httpURLConnection.getInputStream()) {
            int pdfResponseCode = httpURLConnection.getResponseCode();
            long pdfSize = httpURLConnection.getContentLengthLong();

            createCell(row, 5, String.valueOf(pdfSize));

            if (pdfResponseCode == 200 && pdfSize <= 800 * 1024) {
                System.out.println("PDF Verification Passed for URL: " + urlString);
            } else {
                System.out.println("PDF Verification Failed for URL: " + urlString);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void verifyImageResponse(String urlString, HttpURLConnection httpURLConnection, Row row) {
        try (InputStream inputStream = httpURLConnection.getInputStream()) {
            int imageResponseCode = httpURLConnection.getResponseCode();
            long imageSize = httpURLConnection.getContentLengthLong();

            createCell(row, 6, String.valueOf(imageSize));

            if (imageResponseCode == 200 && imageSize <= 800 * 1024) {
                System.out.println("Image Verification Passed for URL: " + urlString);
            } else {
                System.out.println("Image Verification Failed for URL: " + urlString);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void logException(String url, IOException e) {
        System.err.println("Exception occurred for URL: " + url + "\n" + e.getMessage());
    }

    private static boolean isWellFormattedHttpUrl(String url) {
        try {
            URL u = new URL(url);
            return "http".equalsIgnoreCase(u.getProtocol()) || "https".equalsIgnoreCase(u.getProtocol());
        } catch (MalformedURLException e) {
            return false;
        }
    }

    private static boolean isSameDomain(String startingUrl, String url) {
        try {
            URI startingUri = new URI(startingUrl);
            URI uri = new URI(url);
            return Objects.equals(startingUri.getHost(), uri.getHost());
        } catch (URISyntaxException e) {
            e.printStackTrace();
            return false;
        }
    }

    private static void createCell(Row row, int index, String value) {
        Cell cell = row.createCell(index);
        int maxLength = 2000;
        if (value.length() > maxLength) {
            value = value.substring(0, maxLength);
        }
        cell.setCellValue(value);
    }

    private static void logError(String message) {
        System.err.println("ERROR: " + message);
    }

    public static String[] apiResponse(String urlString, String cookie) {
        StringBuilder responseText = new StringBuilder();
        String[] apiResponse = new String[2];

        if (urlString.contains("/logout") || urlString.contains("/admin/flush")) {
            apiResponse[0] = "200";
            apiResponse[1] = "Skipped";
        } else {
            try {
                URL url = new URL(urlString);
                URLConnection connection = url.openConnection();
                connection.setRequestProperty("Cookie", cookie);

                if (connection instanceof HttpURLConnection) {
                    HttpURLConnection httpURLConnection = (HttpURLConnection) connection;

                    httpURLConnection.setRequestMethod("GET");

                    startTime = System.currentTimeMillis();

                    try (InputStream inputStream = httpURLConnection.getInputStream()) {
                        String contentType = httpURLConnection.getContentType();
                        responseText.append(readInputStream(inputStream));
                    } finally {
                        endTime = System.currentTimeMillis();
                    }

                    apiResponse[0] = String.valueOf(httpURLConnection.getResponseCode());
                    apiResponse[1] = responseText.toString();
                } else {
                    apiResponse[0] = "200";
                    apiResponse[1] = "Skipped - Unsupported protocol: " + url.getProtocol();
                }
            } catch (IOException e) {
                if (e instanceof FileNotFoundException || e instanceof UnknownServiceException) {
                    apiResponse[0] = "200";
                    apiResponse[1] = "Skipped - Image/PDF";
                } else {
                    StringWriter sw = new StringWriter();
                    PrintWriter pw = new PrintWriter(sw);
                    e.printStackTrace(pw);
                    String stackTrace = sw.toString();
                    apiResponse[0] = stackTrace;
                }
            }
        }

        return apiResponse;
    }

    private static String readInputStream(InputStream inputStream) throws IOException {
        StringBuilder textBuilder = new StringBuilder();
        try (Reader reader = new BufferedReader(new InputStreamReader(inputStream, Charset.forName(StandardCharsets.UTF_8.name())))) {
            int c;
            while ((c = reader.read()) != -1) {
                textBuilder.append((char) c);
            }
        }
        return textBuilder.toString();
    }

    private static String getResult(String url, String responseCode, String responseText, Sheet resultsSheet) {
        if (!Objects.equals(responseCode, "200")) {
            System.out.println("URL: " + url + ", Verification Failed. Response Code: " + responseCode);
            createErrorRow(url, urlType, "Verification Failed. Response Code:" + responseCode, resultsSheet);
            return "Verification Failed. Response Code:" + responseCode;
        } else if (responseText.contains("Fatal error")) {
            System.out.println("URL: " + url + ", Verification Failed. Response Text: Fatal error");
            createErrorRow(url, urlType, "Verification Failed. Response Text: Fatal error", resultsSheet);
            return "Verification Failed. Response Text: Fatal error";
        } else if (responseText.contains("Page not found")) {
            System.out.println("URL: " + url + ", Verification Failed. Response Text: Page not found");
            createErrorRow(url, urlType, "Verification Failed. Response Text: Page not found", resultsSheet);
            return " Verification Failed. Response Text: Page not found";
        } else if (responseText.contains("error message")) {
            System.out.println("URL: " + url + ", Verification Failed. Response Text: error message");
            createErrorRow(url, urlType, "Verification Failed. Response Text: error message", resultsSheet);
            return " Verification Failed. Response Text: error message";
        } else if (responseText.contains("The website encountered an unexpected error")) {
            System.out.println("URL: " + url + ", Verification Failed. Response Text: The website encountered an unexpected error");
            createErrorRow(url, urlType, "Verification Failed. Response Text: The website encountered an unexpected error", resultsSheet);
            return "failed";
        } else {
            System.out.println("URL: " + url + ", Verification Passed");
            return "Passed";
        }
    }

    private static void createErrorRow(String url, String urlType, String errorDetails, Sheet resultsSheet) {
        Row errorRow = resultsSheet.createRow(resultsSheet.getLastRowNum() + 1);
        createCell(errorRow, 0, url);
        createCell(errorRow, 1, urlType);
        createCell(errorRow, 2, "Verification Failed");
        createCell(errorRow, 3, "N/A");
        createCell(errorRow, 4, String.valueOf(endTime - startTime));
        createCell(errorRow, 5, "N/A");
        createCell(errorRow, 6, "N/A");
        createCell(errorRow, 7, "Error Details: " + errorDetails);
    }
}
