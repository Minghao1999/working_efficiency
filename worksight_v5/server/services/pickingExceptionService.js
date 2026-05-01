const WMS_URL = "https://iwms.us.jdlglobal.com/reportApi/services/smartQueryWS?wsdl";

const WMS_HEADERS = {
  "Content-Type": "text/xml; charset=UTF-8",
  Accept: "application/xml, text/xml, */*; q=0.01",
  Authorization:
    "Bearer eyJhbGciOiJIUzI1NiJ9.eyJkdXJhdGlvbiI6ODY0MDAsImxvZ2luQWNjb3VudCI6ImpkaGtfa0ZHb1V0bWJOU0tGIiwibG9naW5UaW1lIjoiMjAyNi0wMy0yNSAyMzoxNTo1MCIsIm9yZ05vIjoiMSIsImxvZ2luVHlwZSI6ImIiLCJpc0F1dGhJZ25vcmVkIjpmYWxzZSwidGVuYW50Tm8iOiJBMDAwMDAwMDAzMyIsImxvZ2luQ2xpZW50IjoiUEMiLCJkaXN0cmlidXRlTm8iOiIxIiwid2FyZWhvdXNlTm8iOiJDMDAwMDAwOTk0MyIsInRpbWVzdGFtcCI6MTc3NDQ1MTc1MDY4Mn0.iMpyspQLNS47zfOvO2Fq-mpsp-oCoSt1eGZayCrUitM",
};

const WAREHOUSES = {
  "1": { label: "EWR-LG-1-US", warehouseNo: "C0000000389" },
  "2": { label: "EWR-SM-2-US", warehouseNo: "C0000002427" },
  "5": { label: "EWR-LG-5-US", warehouseNo: "C0000009943" }
};

function resolveWarehouse(warehouseKey) {
  return WAREHOUSES[String(warehouseKey || "2")] || WAREHOUSES["2"];
}

function getLast7DaysRange() {
  const now = new Date();
  const start = new Date(now);
  start.setDate(start.getDate() - 7);

  const pad = (value) => String(value).padStart(2, "0");
  const formatDate = (date) =>
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  const formatTime = (date) =>
    `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;

  return {
    startTime: `${formatDate(start)} 00:00:00`,
    endTime: `${formatDate(now)} ${formatTime(now)}`
  };
}

function buildSoapEnvelope(arg0, arg1) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <wms3:queryWs xmlns:wms3="http://wms3.360buy.com">
      <arg0>${JSON.stringify(arg0)}</arg0>
      <arg1>${JSON.stringify(arg1)}</arg1>
    </wms3:queryWs>
  </soap:Body>
</soap:Envelope>`;
}

function buildPickingBody(barcode, containerNo, warehouseNo) {
  const { startTime, endTime } = getLast7DaysRange();
  const arg0 = {
    bizType: "queryReportByCondition",
    uuid: "1",
    callCode: "360BUY.WMS3.WS.CALLCODE.10401"
  };
  const arg1 = {
    Id: "wms_picking_data_v2",
    Name: "pickingResultsOfQuery",
    WkNo: "jdhk_ulXbWlUMYeET",
    UserName: "minghao.sun@jd.com",
    ReportModelId: "",
    SqlLimit: "5000",
    ListSqlOrder: [],
    ListSqlWhere: [
      {
        FieldId: "UPDATE_TIME",
        FieldName: "UPDATE_TIME",
        Compare: 9,
        FirstValue: startTime,
        SecondValue: endTime,
        Location: ""
      },
      {
        FieldId: "barcode",
        FieldName: "barcode",
        Compare: 0,
        FirstValue: barcode,
        Location: ""
      },
      {
        FieldId: "pick_container_no",
        FieldName: "pick_container_no",
        Compare: 0,
        FirstValue: containerNo,
        Location: ""
      }
    ],
    PageSize: 100,
    CurrentPage: 1,
    orgNo: "1",
    distributeNo: "1",
    warehouseNo
  };

  return buildSoapEnvelope(arg0, arg1);
}

function buildInventoryBody(barcode, warehouseNo) {
  const arg0 = {
    bizType: "queryReportByCondition",
    uuid: "1",
    callCode: "360BUY.WMS3.WS.CALLCODE.10401"
  };
  const arg1 = {
    Id: "stockReport",
    Name: "commodityInventoryInformationInquiry",
    WkNo: "jdhk_kFGoUtmbNSKF",
    UserName: "wujianghao0706@gmail.com",
    ReportModelId: "",
    SqlLimit: "5000",
    ListSqlOrder: [],
    ListSqlWhere: [
      {
        FieldId: "barcode",
        FieldName: "barcode",
        Compare: 0,
        FirstValue: barcode,
        Location: ""
      }
    ],
    PageSize: 100,
    CurrentPage: 1,
    orgNo: "1",
    distributeNo: "1",
    warehouseNo
  };

  return buildSoapEnvelope(arg0, arg1);
}

function decodeXmlText(value) {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function parseResponseText(xmlText) {
  const decoded = decodeXmlText(xmlText);
  const start = decoded.indexOf("{");
  const end = decoded.lastIndexOf("}");

  if (start === -1 || end === -1 || end <= start) {
    return {};
  }

  try {
    return JSON.parse(decoded.slice(start, end + 1));
  } catch {
    return {};
  }
}

function getRows(result) {
  let rows = result.data || result.rows || result.list || result.result || [];

  if (typeof rows === "string") {
    try {
      rows = JSON.parse(rows);
    } catch {
      rows = [];
    }
  }

  return Array.isArray(rows) ? rows : [];
}

async function postWms(body, warehouseNo) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);

  try {
    const response = await fetch(WMS_URL, {
      method: "POST",
      headers: {
        ...WMS_HEADERS,
        routerule: `1,1,${warehouseNo}`
      },
      body,
      signal: controller.signal
    });

    if (!response.ok) {
      throw new Error(`WMS request failed with ${response.status}`);
    }

    return parseResponseText(await response.text());
  } finally {
    clearTimeout(timeout);
  }
}

async function queryLocationByBarcodeAndContainer(barcode, containerNo, warehouseNo) {
  const result = await postWms(buildPickingBody(barcode, containerNo, warehouseNo), warehouseNo);

  if (!Object.keys(result).length) {
    return {
      success: false,
      message: "Picking API returned no data or could not be parsed.",
      location: null,
      raw: null
    };
  }

  const rows = getRows(result);

  if (!rows.length) {
    return {
      success: false,
      message: "No data found.",
      location: null,
      raw: result
    };
  }

  const possibleLocationFields = [
    "location_no",
    "locationCode",
    "cellNo",
    "cell_code",
    "goods_location",
    "recommondLocation",
    "target_location"
  ];

  const locations = rows
    .map((row) => {
      const field = possibleLocationFields.find((key) => row[key]);
      return field ? String(row[field]).trim().toUpperCase() : "";
    })
    .filter(Boolean);

  return {
    success: true,
    message: "Query succeeded.",
    locations: [...new Set(locations)],
    raw: rows
  };
}

async function queryInventoryByBarcode(barcode, warehouseNo) {
  const result = await postWms(buildInventoryBody(barcode, warehouseNo), warehouseNo);
  const inventoryMap = {};

  for (const row of getRows(result)) {
    const location = row.cellNo || row.locationCode || row.goods_location;
    const qty = Number(row.cellQty ?? 0);

    if (location) {
      inventoryMap[String(location).trim().toUpperCase()] = Number.isFinite(qty) ? qty : 0;
    }
  }

  return inventoryMap;
}

export async function searchPickingExceptionLocation({ barcode, containerNo, warehouse }) {
  const cleanBarcode = String(barcode || "").trim();
  const cleanContainerNo = String(containerNo || "").trim();
  const selectedWarehouse = resolveWarehouse(warehouse);
  const logs = [
    "==================== NEW REQUEST ====================",
    `warehouse: ${selectedWarehouse.label} (${selectedWarehouse.warehouseNo})`,
    `barcode: ${cleanBarcode}`,
    `container: ${cleanContainerNo}`
  ];

  if (!cleanBarcode) {
    return { success: false, message: "Product barcode is required." };
  }

  if (!cleanContainerNo) {
    return { success: false, message: "Container number is required." };
  }

  try {
    logs.push("[STEP 1] Calling picking API...");
    const pickResult = await queryLocationByBarcodeAndContainer(
      cleanBarcode,
      cleanContainerNo,
      selectedWarehouse.warehouseNo
    );

    if (!pickResult.success) {
      return { ...pickResult, logs };
    }

    const pickLocations = pickResult.locations || [];
    logs.push(`Picking locations: ${pickLocations.length ? pickLocations.join(", ") : "(empty)"}`);

    if (!pickLocations.length) {
      logs.push("No picking locations found.");
      return { success: false, message: "No picking locations found.", logs };
    }

    logs.push("[STEP 2] Calling inventory API...");
    const inventoryMap = await queryInventoryByBarcode(cleanBarcode, selectedWarehouse.warehouseNo);
    logs.push("Inventory API returned:");
    for (const [location, qty] of Object.entries(inventoryMap)) {
      logs.push(`  ${location} -> ${qty}`);
    }

    if (!Object.keys(inventoryMap).length) {
      logs.push("Inventory API returned no data.");
      return { success: false, message: "Inventory API returned no data.", logs };
    }

    logs.push("[STEP 3] Calculating intersection...");
    const candidates = pickLocations
      .map((location) => {
        const cleanLocation = String(location).trim().toUpperCase();
        if (Object.prototype.hasOwnProperty.call(inventoryMap, cleanLocation)) {
          logs.push(`Hit: ${cleanLocation} -> inventory ${inventoryMap[cleanLocation]}`);
          return { location: cleanLocation, qty: inventoryMap[cleanLocation] };
        }

        logs.push(`Miss inventory: ${cleanLocation}`);
        return null;
      })
      .filter(Boolean);

    if (!candidates.length) {
      logs.push("No intersection. Using fallback logic: choose the lowest inventory location.");
      const fallbackList = Object.entries(inventoryMap).map(([location, qty]) => ({ location, qty }));
      const bestFallback = fallbackList.sort((left, right) => left.qty - right.qty)[0];

      if (!bestFallback) {
        logs.push("Inventory list is empty; no recommendation available.");
        return { success: false, message: "Inventory API is empty; no recommendation available.", logs };
      }

      logs.push("Fallback candidates:");
      for (const item of fallbackList) {
        logs.push(`  ${item.location} -> ${item.qty}`);
      }
      logs.push(`Fallback recommendation: ${bestFallback.location} (inventory ${bestFallback.qty})`);
      logs.push("====================================================");

      return {
        success: true,
        location: bestFallback.location,
        qty: bestFallback.qty,
        type: "fallback_inventory",
        candidates: [],
        pickLocations,
        inventoryCount: fallbackList.length,
        warehouse: selectedWarehouse,
        logs
      };
    }

    logs.push("[STEP 4] Choosing the lowest inventory location...");
    logs.push("Candidates:");
    for (const candidate of candidates) {
      logs.push(`  ${candidate.location} -> ${candidate.qty}`);
    }
    const best = [...candidates].sort((left, right) => left.qty - right.qty)[0];
    logs.push(`Final recommendation: ${best.location} (inventory ${best.qty})`);
    logs.push("====================================================");

    return {
      success: true,
      location: best.location,
      qty: best.qty,
      type: "matched_pick_inventory",
      candidates,
      pickLocations,
      inventoryCount: Object.keys(inventoryMap).length,
      warehouse: selectedWarehouse,
      logs
    };
  } catch (error) {
    logs.push(`Exception: ${error?.message || error}`);
    if (String(error?.message || error).includes("NoneType")) {
      logs.push("Inventory API exception. Returning extra.");
      return {
        success: true,
        location: "extra",
        type: "inventory_error",
        logs
      };
    }

    return {
      success: false,
      message: error?.name === "AbortError" ? "WMS request timed out." : error.message,
      logs
    };
  }
}
