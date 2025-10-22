private static JsonObject getTopAgencyByEstimate(JsonObject totalData) {
    if (totalData == null || !totalData.has("data") || totalData.get("data").isJsonNull()) {
        System.out.println("⚠️ No 'data' found in totalData response.");
        return null;
    }

    JsonElement dataElem = totalData.get("data");
    if (!dataElem.isJsonArray()) {
        System.out.println("⚠️ 'data' is not a JsonArray.");
        return null;
    }

    JsonArray dataArray = dataElem.getAsJsonArray();
    JsonObject topAgency = null;
    double maxEstimate = -1;

    for (JsonElement elem : dataArray) {
        JsonObject record = elem.getAsJsonObject();
        double estimateTotal = 0;
        if (record.has("estimateTotal") && !record.get("estimateTotal").isJsonNull()) {
            try {
                estimateTotal = record.get("estimateTotal").getAsDouble();
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
