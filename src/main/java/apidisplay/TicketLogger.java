package apidisplay;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;

public class TicketLogger {

	private static final String TOTAL_EST_URL = System.getenv("TOTAL_EST_URL");
	private static final String LAST_WEEK_URL = System.getenv("LAST_WEEK_URL");
	private static final String SLACK_WEBHOOK_URL = System.getenv("SLACK_WEBHOOK_URL");
  // replace with your webhook

    private static final Gson gson = new Gson();

    public static void main(String[] args) {
        try {
            JsonObject totalEstResp = fetchJson(TOTAL_EST_URL);
            JsonObject lastWeekResp = fetchJson(LAST_WEEK_URL);

            String slackText = buildSlackTable(totalEstResp, lastWeekResp);

            postToSlack(SLACK_WEBHOOK_URL, slackText);

            System.out.println("Posted to Slack successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static JsonObject fetchJson(String url) throws Exception {
        HttpClient client = HttpClient.newHttpClient();
        HttpRequest req = HttpRequest.newBuilder()
            .uri(URI.create(url))
            .GET()
            .build();
        HttpResponse<String> resp = client.send(req, HttpResponse.BodyHandlers.ofString());

        return gson.fromJson(resp.body(), JsonObject.class);
    }

    private static String buildSlackTable(JsonObject totalResp, JsonObject lastWeekResp) {
        StringBuilder sb = new StringBuilder();
        sb.append("*Usage Report*\n");

        // Use code block for monospaced font
        sb.append("```");
        // Header with aligned columns
        sb.append(String.format("%-22s | %-12s | %-10s | %-20s\n",
                "Agency", "Branch", "Medium", "EstTotal/RO/IB/OB"));
        sb.append(String.format("%-22s-+-%-12s-+-%-10s-+-%-20s\n",
                repeatChar('-', 22), repeatChar('-', 12), repeatChar('-', 10), repeatChar('-', 20)));


        appendRows(sb, totalResp, "Usage Report");
        appendRows(sb, lastWeekResp, "Last Week Usage Report");

        sb.append("```");
        return sb.toString();
    }

    private static void appendRows(StringBuilder sb, JsonObject respObj, String sectionLabel) {
        if (respObj == null) return;

        sb.append("\n").append(sectionLabel).append(":\n");
        JsonArray arr = respObj.getAsJsonArray("data");
        if (arr == null) return;

        for (JsonElement je : arr) {
            JsonObject rec = je.getAsJsonObject();
            String agency = rec.has("agencyName") && !rec.get("agencyName").isJsonNull()
                    ? trimTo(rec.get("agencyName").getAsString(), 20)
                    : "null";
            String branch = rec.has("branchName") && !rec.get("branchName").isJsonNull()
                    ? trimTo(rec.get("branchName").getAsString(), 8)
                    : "null";
            String medium = rec.has("medium") && !rec.get("medium").isJsonNull()
                    ? trimTo(rec.get("medium").getAsString(), 8)
                    : "null";

            String estTotal = rec.has("estimateTotal") && !rec.get("estimateTotal").isJsonNull()
                    ? rec.get("estimateTotal").getAsString()
                    : "null";
            String ro = rec.has("roTotal") && !rec.get("roTotal").isJsonNull()
                    ? rec.get("roTotal").getAsString()
                    : "null";
            String ib = rec.has("ibTotal") && !rec.get("ibTotal").isJsonNull()
                    ? rec.get("ibTotal").getAsString()
                    : "null";
            String ob = rec.has("obTotal") && !rec.get("obTotal").isJsonNull()
                    ? rec.get("obTotal").getAsString()
                    : "null";

            String combined = String.format("%s/%s/%s/%s", estTotal, ro, ib, ob);

            sb.append(String.format("%-22s | %-12s | %-10s | %-20s\n",
                    trimTo(agency, 22),
                    trimTo(branch, 12),
                    trimTo(medium, 10),
                    combined));

        }
    }

    private static String trimTo(String s, int maxLen) {
        if (s == null) return "null";
        if (s.length() >= maxLen) return s.substring(0, maxLen);
        return String.format("%-" + maxLen + "s", s);  // pad with spaces
    }


    private static String repeatChar(char c, int count) {
        return new String(new char[count]).replace('\0', c);
    }

    private static void postToSlack(String webhookUrl, String text) throws Exception {
        JsonObject payload = new JsonObject();
        payload.addProperty("text", text);

        HttpClient client = HttpClient.newHttpClient();
        HttpRequest req = HttpRequest.newBuilder()
            .uri(URI.create(webhookUrl))
            .header("Content-Type", "application/json")
            .POST(HttpRequest.BodyPublishers.ofString(gson.toJson(payload)))
            .build();
        HttpResponse<String> resp = client.send(req, HttpResponse.BodyHandlers.ofString());

        System.out.println("Slack API response: " + resp.statusCode() + " — " + resp.body());
    }
}
