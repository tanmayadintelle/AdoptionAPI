package apidisplay;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;

import com.slack.api.Slack;
import com.slack.api.methods.SlackApiException;
import com.slack.api.methods.request.files.FilesUploadV2Request;
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
	    private static final String SLACK_BOT_TOKEN = System.getenv("SLACK_BOT_TOKEN");  // Replace
	    private static final String SLACK_CHANNEL_ID = System.getenv("SLACK_CHANNEL_ID");;         // Replace

    private static final Gson gson = new Gson();
    private static final HttpClient httpClient = HttpClient.newHttpClient();

    public static void main(String[] args) {
        try {
            JsonObject total = fetchJson(TOTAL_EST_URL);
            JsonObject lastWeek = fetchJson(LAST_WEEK_URL);

            byte[] excelBytes = createExcelFile(total, lastWeek);

            uploadFileV2(excelBytes, "UsageReport.xlsx");

            System.out.println("File upload via V2 succeeded.");
        } catch (IOException | SlackApiException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static JsonObject fetchJson(String url) throws Exception {
        HttpRequest req = HttpRequest.newBuilder()
            .uri(URI.create(url))
            .GET()
            .build();
        HttpResponse<String> resp = httpClient.send(req, HttpResponse.BodyHandlers.ofString());
        return gson.fromJson(resp.body(), JsonObject.class);
    }

    private static byte[] createExcelFile(JsonObject total, JsonObject lastWeek) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Usage Report");

        // Header
        Row header = sheet.createRow(0);
        String[] headers = { "Agency", "Branch", "Medium", "EstTotal", "RO", "IB", "OB", "Report Type" };
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
        if (respObj == null) return startRow;
        JsonArray arr = respObj.getAsJsonArray("data");
        if (arr == null) return startRow;

        int rowNum = startRow;
        for (JsonElement je : arr) {
            JsonObject rec = je.getAsJsonObject();
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(getSafeString(rec, "agencyName"));
            row.createCell(1).setCellValue(getSafeString(rec, "branchName"));
            row.createCell(2).setCellValue(getSafeString(rec, "medium"));
            row.createCell(3).setCellValue(getSafeString(rec, "estimateTotal"));
            row.createCell(4).setCellValue(getSafeString(rec, "roTotal"));
            row.createCell(5).setCellValue(getSafeString(rec, "ibTotal"));
            row.createCell(6).setCellValue(getSafeString(rec, "obTotal"));
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
                .initialComment("Here is the latest usage report")
                .title("UsageReport")
                .build();

        FilesUploadV2Response resp = slack.methods().filesUploadV2(req);

        if (!resp.isOk()) {
            throw new RuntimeException("Upload V2 failed: " + resp.getError());
        }
        System.out.println("V2 upload OK, file data: " + resp.getFiles());
    }
}
