package apidisplay;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.*;

import com.slack.api.Slack;
import com.slack.api.methods.SlackApiException;
import com.slack.api.methods.request.chat.ChatPostMessageRequest;
import com.slack.api.methods.request.files.FilesUploadV2Request;
import com.slack.api.methods.response.chat.ChatPostMessageResponse;
import com.slack.api.methods.response.files.FilesUploadV2Response;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.*;

public class TicketLogger {

    private static final String TOTAL_EST_URL = System.getenv("TOTAL_EST_URL");
    private static final String LAST_WEEK_URL = System.getenv("LAST_WEEK_URL");
    private static final String SLACK_BOT_TOKEN = System.getenv("SLACK_BOT_TOKEN");
    private static final String SLACK_CHANNEL_ID = System.getenv("SLACK_CHANNEL_ID");

    private static final Gson gson = new Gson();
    private static final HttpClient httpClient = HttpClient.newHttpClient();

    // ---------------- MAIN ----------------
    public static void main(String[] args) {
        try {
            JsonObject total = fetchJson(TOTAL_EST_URL);
            JsonObject lastWeek = fetchJson(LAST_WEEK_URL);

            byte[] excelBytes = createExcelFile(total, lastWeek);
            uploadFileV2(excelBytes, "UsageReport.xlsx");

            List<AgencyEstimate> sorted = getSortedAgenciesByEstimate(total);

            if (sorted.size() >= 5) {
                List<AgencyEstimate> top5 = sorted.subList(0, 5);
                List<AgencyEstimate> bottom5 =
                        sorted.subList(Math.max(sorted.size() - 5, 0), sorted.size());

                String message =
                        "📊 *Estimate Usage Summary*\n\n" +
                        formatAgencyList("🏆 *Top 5 Agencies (Highest Usage)*", top5) +
                        "\n" +
                        formatAgencyList("⚠️ *Bottom 5 Agencies (Lowest Usage)*", bottom5);

                sendSlackMessage(message);
            } else {
                sendSlackMessage("⚠️ Not enough data to calculate Top 5 / Bottom 5 agencies.");
            }

            System.out.println("✅ Report uploaded & Slack message sent.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ---------------- API ----------------
    private static JsonObject fetchJson(String url) throws Exception {
        HttpRequest req = HttpRequest.newBuilder()
                .uri(URI.create(url))
                .header("Content-Type", "application/json")
                .POST(HttpRequest.BodyPublishers.ofString("{}"))
                .build();

        HttpResponse<String> resp = httpClient.send(req, HttpResponse.BodyHandlers.ofString());

        if (resp.statusCode() != 200) {
            throw new IOException("HTTP error: " + resp.statusCode());
        }
        return gson.fromJson(resp.body(), JsonObject.class);
    }

    // ---------------- EXCEL ----------------
    private static byte[] createExcelFile(JsonObject total, JsonObject lastWeek) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Usage Report");

        String[] headers = { "Agency", "Branch", "Medium", "EstTotal", "RO", "IB", "OB", "Report Type" };
        Row header = sheet.createRow(0);

        for (int i = 0; i < headers.length; i++) {
            header.createCell(i).setCellValue(headers[i]);
        }

        int rowNum = 1;
        rowNum = writeRows(sheet, total, rowNum, "Total Estimates");
        writeRows(sheet, lastWeek, rowNum, "Last Week Estimates");

        for (int i = 0; i < headers.length; i++) sheet.autoSizeColumn(i);

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        workbook.close();
        return bos.toByteArray();
    }

    private static int writeRows(Sheet sheet, JsonObject obj, int rowNum, String type) {
        if (obj == null || !obj.has("Data")) return rowNum;

        JsonArray table = obj.getAsJsonObject("Data").getAsJsonArray("Table");
        if (table == null) return rowNum;

        for (JsonElement e : table) {
            JsonObject r = e.getAsJsonObject();
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(getSafe(r, "Agency Name"));
            row.createCell(1).setCellValue(getSafe(r, "BranchName"));
            row.createCell(2).setCellValue(getSafe(r, "Medium"));
            row.createCell(3).setCellValue(getSafe(r, "Total Estimate"));
            row.createCell(4).setCellValue(getSafe(r, "Total RO"));
            row.createCell(5).setCellValue(getSafe(r, "Total IB"));
            row.createCell(6).setCellValue(getSafe(r, "Total OB"));
            row.createCell(7).setCellValue(type);
        }
        return rowNum;
    }

    // ---------------- TOP/BOTTOM ----------------
    static class AgencyEstimate {
        String agency;
        String branch;
        double estimate;

        AgencyEstimate(String agency, String branch, double estimate) {
            this.agency = agency;
            this.branch = branch;
            this.estimate = estimate;
        }
    }

    private static List<AgencyEstimate> getSortedAgenciesByEstimate(JsonObject total) {
        List<AgencyEstimate> list = new ArrayList<>();

        if (total == null || !total.has("Data")) return list;
        JsonArray table = total.getAsJsonObject("Data").getAsJsonArray("Table");
        if (table == null) return list;

        for (JsonElement e : table) {
            JsonObject r = e.getAsJsonObject();

            String agency = getSafe(r, "Agency Name");
            String branch = getSafe(r, "BranchName");
            double est = 0;

            try {
                est = r.get("Total Estimate").getAsDouble();
            } catch (Exception ignored) {}

            list.add(new AgencyEstimate(agency, branch, est));
        }

        list.sort((a, b) -> Double.compare(b.estimate, a.estimate));
        return list;
    }

    private static String formatAgencyList(String title, List<AgencyEstimate> list) {
        StringBuilder sb = new StringBuilder(title).append("\n");
        int i = 1;

        for (AgencyEstimate a : list) {
            sb.append(i++).append(". ")
              .append(a.agency)
              .append(" (")
              .append(a.branch)
              .append(")")
              .append(" → ")
              .append(a.estimate)
              .append("\n");
        }
        return sb.toString();
    }

    // ---------------- SLACK ----------------
    private static void uploadFileV2(byte[] bytes, String filename)
            throws IOException, SlackApiException {

        Slack slack = Slack.getInstance();
        FilesUploadV2Request req = FilesUploadV2Request.builder()
                .token(SLACK_BOT_TOKEN)
                .channel(SLACK_CHANNEL_ID)
                .filename(filename)
                .fileData(bytes)
                .initialComment("📊 Latest Usage Report")
                .build();

        FilesUploadV2Response resp = slack.methods().filesUploadV2(req);
        if (!resp.isOk()) throw new RuntimeException(resp.getError());
    }

    private static void sendSlackMessage(String msg)
            throws IOException, SlackApiException {

        Slack slack = Slack.getInstance();
        ChatPostMessageResponse resp =
                slack.methods().chatPostMessage(ChatPostMessageRequest.builder()
                        .token(SLACK_BOT_TOKEN)
                        .channel(SLACK_CHANNEL_ID)
                        .text(msg)
                        .build());

        if (!resp.isOk()) throw new RuntimeException(resp.getError());
    }

    // ---------------- UTIL ----------------
    private static String getSafe(JsonObject o, String k) {
        return o.has(k) && !o.get(k).isJsonNull() ? o.get(k).getAsString() : "";
    }
}
