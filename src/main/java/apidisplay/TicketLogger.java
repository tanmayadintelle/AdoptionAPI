	package apidisplay;
	
	import java.io.ByteArrayOutputStream;
	import java.io.IOException;
	import java.net.URI;
	import java.net.http.HttpClient;
	import java.net.http.HttpRequest;
	import java.net.http.HttpResponse;
	
	import com.slack.api.Slack;
	import com.slack.api.methods.SlackApiException;
	import com.slack.api.methods.request.chat.ChatPostMessageRequest;
	import com.slack.api.methods.request.files.FilesUploadV2Request;
	import com.slack.api.methods.response.chat.ChatPostMessageResponse;
	import com.slack.api.methods.response.files.FilesUploadV2Response;
	
	import org.apache.poi.ss.usermodel.*;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	
	import com.google.gson.Gson;
	import com.google.gson.JsonArray;
	import com.google.gson.JsonElement;
	import com.google.gson.JsonObject;
	
	public class TicketLogger {
	
	    private static final String TOTAL_EST_URL = System.getenv("TOTAL_EST_URL");
	    private static final String LAST_WEEK_URL = System.getenv("LAST_WEEK_URL");
	    private static final String SLACK_BOT_TOKEN = System.getenv("SLACK_BOT_TOKEN");
	    private static final String SLACK_CHANNEL_ID = System.getenv("SLACK_CHANNEL_ID");
	
	    private static final Gson gson = new Gson();
	    private static final HttpClient httpClient = HttpClient.newHttpClient();
	
	    public static void main(String[] args) {
	        try {
	            JsonObject total = fetchJson(TOTAL_EST_URL);
	            JsonObject lastWeek = fetchJson(LAST_WEEK_URL);
	
	            byte[] excelBytes = createExcelFile(total, lastWeek);
	            uploadFileV2(excelBytes, "UsageReport.xlsx");
	
	            JsonObject topAgency = getTopAgencyByEstimate(total);
	            if (topAgency != null) {
	                String agency = topAgency.has("agencyName") ? topAgency.get("agencyName").getAsString() : "Unknown";
	                String estimate = topAgency.has("estimateTotal") ? topAgency.get("estimateTotal").getAsString() : "0";
	                sendSlackMessage("🏆 *Top Agency by Estimate Total*: " + agency + " with estimate total of *" + estimate + "*.");
	            } else {
	                sendSlackMessage("⚠️ Could not determine top agency — data might be missing or malformed.");
	            }
	
	            System.out.println("File upload and message sent successfully.");
	        } catch (IOException | SlackApiException e) {
	            e.printStackTrace();
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	    }
	
	    private static JsonObject fetchJson(String url) throws Exception {
	        HttpRequest req = HttpRequest.newBuilder()
	                .uri(URI.create(url))
	                .header("Content-Type", "application/json")
	                // Send an empty JSON body or modify if the API expects parameters
	                .POST(HttpRequest.BodyPublishers.ofString("{}"))
	                .build();
	        HttpResponse<String> resp = httpClient.send(req, HttpResponse.BodyHandlers.ofString());
	        if (resp.statusCode() != 200) {
	            throw new IOException("HTTP error: " + resp.statusCode() + " - " + resp.body());
	        }
	
	        return gson.fromJson(resp.body(), JsonObject.class);
	    }
	
	    private static byte[] createExcelFile(JsonObject total, JsonObject lastWeek) throws IOException {
	        Workbook workbook = new XSSFWorkbook();
	        Sheet sheet = workbook.createSheet("Usage Report");
	
	        String[] headers = { "Agency", "Branch", "Medium", "EstTotal", "RO", "IB", "OB", "Report Type" };
	        Row header = sheet.createRow(0);
	        for (int i = 0; i < headers.length; i++) {
	            header.createCell(i).setCellValue(headers[i]);
	        }
	
	        int rowNum = 1;
	        rowNum = writeRowsToSheet(sheet, total, rowNum, "Total Estimates");
	        rowNum = writeRowsToSheet(sheet, lastWeek, rowNum, "Last Week Estimates");
	
	        for (int i = 0; i < headers.length; i++) {
	            sheet.autoSizeColumn(i);
	        }
	
	        ByteArrayOutputStream bos = new ByteArrayOutputStream();
	        workbook.write(bos);
	        workbook.close();
	        return bos.toByteArray();
	    }
	
	    private static int writeRowsToSheet(Sheet sheet, JsonObject respObj, int startRow, String reportType) {
	        if (respObj == null || !respObj.has("Data")) return startRow;
	
	        JsonObject dataObj = respObj.getAsJsonObject("Data");
	        if (!dataObj.has("Table") || !dataObj.get("Table").isJsonArray()) return startRow;
	
	        JsonArray arr = dataObj.getAsJsonArray("Table");
	        int rowNum = startRow;
	        for (JsonElement je : arr) {
	            JsonObject rec = je.getAsJsonObject();
	            Row row = sheet.createRow(rowNum++);
	            row.createCell(0).setCellValue(getSafeString(rec, "Agency Name"));
	            row.createCell(1).setCellValue(getSafeString(rec, "BranchName"));
	            row.createCell(2).setCellValue(getSafeString(rec, "Medium"));
	            row.createCell(3).setCellValue(getSafeString(rec, "Total Estimate"));
	            row.createCell(4).setCellValue(getSafeString(rec, "Total RO"));
	            row.createCell(5).setCellValue(getSafeString(rec, "Total IB"));
	            row.createCell(6).setCellValue(getSafeString(rec, "Total OB"));
	            row.createCell(7).setCellValue(reportType);
	        }
	        return rowNum;
	    }
	
	    private static String getSafeString(JsonObject obj, String key) {
	        return obj.has(key) && !obj.get(key).isJsonNull() ? obj.get(key).getAsString() : "";
	    }
	
	    private static void uploadFileV2(byte[] fileBytes, String filename) throws IOException, SlackApiException {
	        Slack slack = Slack.getInstance();
	
	        FilesUploadV2Request req = FilesUploadV2Request.builder()
	                .token(SLACK_BOT_TOKEN)
	                .channel(SLACK_CHANNEL_ID)
	                .filename(filename)
	                .fileData(fileBytes)
	                .initialComment("📊 Here is the latest usage report")
	                .title("UsageReport")
	                .build();
	
	        FilesUploadV2Response resp = slack.methods().filesUploadV2(req);
	
	        if (!resp.isOk()) {
	            throw new RuntimeException("Upload V2 failed: " + resp.getError());
	        }
	    }
	
	    private static void sendSlackMessage(String message) throws IOException, SlackApiException {
	        Slack slack = Slack.getInstance();
	        ChatPostMessageRequest request = ChatPostMessageRequest.builder()
	                .token(SLACK_BOT_TOKEN)
	                .channel(SLACK_CHANNEL_ID)
	                .text(message)
	                .build();
	
	        ChatPostMessageResponse response = slack.methods().chatPostMessage(request);
	        if (!response.isOk()) {
	            throw new RuntimeException("Slack message failed: " + response.getError());
	        }
	    }
	
	    private static JsonObject getTopAgencyByEstimate(JsonObject totalData) {
	    	 if (totalData == null || !totalData.has("Data")) {
	    	        System.out.println("⚠️ No 'Data' found in totalData response.");
	    	        return null;
	    	    }
	
	    	    JsonObject dataObj = totalData.getAsJsonObject("Data");
	    	    if (!dataObj.has("Table") || !dataObj.get("Table").isJsonArray()) {
	    	        System.out.println("⚠️ 'Table' is missing or not an array.");
	    	        return null;
	    	    }
	
	    	    JsonArray table = dataObj.getAsJsonArray("Table");
	    	    JsonObject topAgency = null;
	    	    double maxEstimate = -1;
	
	    	    for (JsonElement elem : table) {
	    	        JsonObject record = elem.getAsJsonObject();
	    	        double estimateTotal = 0;
	
	    	        if (record.has("Total Estimate") && !record.get("Total Estimate").isJsonNull()) {
	    	            try {
	    	                estimateTotal = record.get("Total Estimate").getAsDouble();
	    	            } catch (NumberFormatException e) {
	    	                continue;
	    	            }
	    	        }
	
	    	        if (estimateTotal > maxEstimate) {
	    	            maxEstimate = estimateTotal;
	    	            topAgency = record;
	    	        }
	    	    }
	
	    	    return topAgency;
	    	}
	}
