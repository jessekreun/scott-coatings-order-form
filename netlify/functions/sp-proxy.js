// sp-proxy.js — SharePoint proxy for Scott Coatings Order Form
// Proxies all SharePoint API calls server-side to avoid CORS restrictions.
// The browser passes its MSAL token; this function forwards it to SharePoint.

const SITE = "https://scottcoatings.sharepoint.com/sites/ProjectOperations";
const MPI_LIST = "Master Project Index";
const ITEMS_LIST = "SUN";
const ORDERS_LIST = "Project Orders";

const ACTIVE_STATUSES = ["Course of Construction", "Pre-Construction"];

exports.handler = async function(event) {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Authorization, Content-Type",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Content-Type": "application/json"
  };

  // Handle CORS preflight
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers, body: "" };
  }

  const action = event.queryStringParameters?.action;
  const token = event.headers?.authorization;

  if (!token) {
    return {
      statusCode: 401,
      headers,
      body: JSON.stringify({ error: "No authorization token provided" })
    };
  }

  if (!action) {
    return {
      statusCode: 400,
      headers,
      body: JSON.stringify({ error: "No action specified" })
    };
  }

  try {
    switch (action) {
      case "getMPI":
        return await getMPI(token, headers);
      case "getItems":
        return await getItems(token, headers);
      case "submitOrder":
        return await submitOrder(token, headers, event.body);
      default:
        return {
          statusCode: 400,
          headers,
          body: JSON.stringify({ error: `Unknown action: ${action}` })
        };
    }
  } catch (err) {
    console.error(`Error in action ${action}:`, err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: err.message || "Internal server error" })
    };
  }
};

// ── SP FETCH HELPER ──────────────────────────────────────────────────────────

async function spGet(url, token) {
  const res = await fetch(url, {
    headers: {
      "Accept": "application/json;odata=verbose",
      "Authorization": token
    }
  });
  if (!res.ok) {
    const text = await res.text().catch(() => res.statusText);
    throw new Error(`SharePoint ${res.status}: ${text}`);
  }
  const data = await res.json();
  return data.d?.results ?? data.d ?? data;
}

async function spGetAll(baseUrl, token) {
  let results = [];
  let url = baseUrl;
  while (url) {
    const res = await fetch(url, {
      headers: {
        "Accept": "application/json;odata=verbose",
        "Authorization": token
      }
    });
    if (!res.ok) {
      const text = await res.text().catch(() => res.statusText);
      throw new Error(`SharePoint ${res.status}: ${text}`);
    }
    const data = await res.json();
    results = results.concat(data.d?.results ?? []);
    url = data.d?.__next ?? null;
  }
  return results;
}

// ── GET MPI ──────────────────────────────────────────────────────────────────

async function getMPI(token, headers) {
  const statusFilter = ACTIVE_STATUSES
    .map(s => `Project_x0020_Status eq '${encodeURIComponent(s)}'`)
    .join(" or ");

  const url =
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(MPI_LIST)}')/items` +
    `?$select=Id,Title,Project_x0020_Name,Project_x0020_Location,PM_x0020_Name,Client_x0020_Company/Title,Project_x0020_Status` +
    `&$expand=Client_x0020_Company` +
    `&$filter=${statusFilter}` +
    `&$orderby=Project_x0020_Name` +
    `&$top=500`;

  const results = await spGetAll(url, token);

  const jobs = results
    .map(r => ({
      id: r.Id,
      name: r.Project_x0020_Name || "",
      number: r.Title || "",
      address: r.Project_x0020_Location || "",
      pm: r.PM_x0020_Name || "",
      client: r.Client_x0020_Company?.Title || "",
      status: r.Project_x0020_Status || ""
    }))
    .filter(j => j.name);

  return {
    statusCode: 200,
    headers,
    body: JSON.stringify({ jobs })
  };
}

// ── GET ITEMS ────────────────────────────────────────────────────────────────

const SP_SITE_ID = "6b5a71b8-6a22-4cbc-80c1-83281dbce1af,72c919df-0ae4-4b80-85e5-fc3820e7b1e2";
const SP_LIST_ID = "fab56298-a6c9-4aec-befa-7349e41a2aad";

async function getItems(token, headers) {
  const url =
    `${SITE}/_api/web/GetList(@listUrl)/items` +
    `?@listUrl='${encodeURIComponent("/sites/ProjectOperations/Lists/SUN")}'` +
    `&$select=Id,Title,Description,SageID,Manufacturer,VendorID,Packaging` +
    `&$top=500`;

  const results = await spGetAll(url, token);

  const items = results.map(r => ({
    id: r.Id,
    sageId: r.Title || "",
    name: r.Description || r.Title || "Unnamed item",
    type: r.SageID || "",
    mfr: r.Manufacturer || "",
    vendorId: r.VendorID || "",
    pkg: r.Packaging || "",
    picture: `https://scottcoatings.sharepoint.com/_api/v2.1/sites('${SP_SITE_ID}')/lists('${SP_LIST_ID}')/items('${r.Id}')/thumbnails/0/large/content?prefer=noredirect`
  }));

  return {
    statusCode: 200,
    headers,
    body: JSON.stringify({ items })
  };
}

// ── SUBMIT ORDER ─────────────────────────────────────────────────────────────

async function submitOrder(token, headers, rawBody) {
  if (!rawBody) throw new Error("No body provided for submitOrder");

  const payload = JSON.parse(rawBody);

  // Step 1: Get list metadata (entity type name)
  const metaUrl =
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(ORDERS_LIST)}')` +
    `?$select=ListItemEntityTypeFullName`;

  const metaRes = await fetch(metaUrl, {
    headers: {
      "Accept": "application/json;odata=verbose",
      "Authorization": token
    }
  });
  if (!metaRes.ok) throw new Error(`Meta fetch failed: ${metaRes.status}`);
  const metaData = await metaRes.json();
  const entityType = metaData.d.ListItemEntityTypeFullName;

  // Step 2: Get form digest
  const digestRes = await fetch(`${SITE}/_api/contextinfo`, {
    method: "POST",
    headers: {
      "Accept": "application/json;odata=verbose",
      "Authorization": token
    }
  });
  if (!digestRes.ok) throw new Error(`Digest fetch failed: ${digestRes.status}`);
  const digestData = await digestRes.json();
  const digest = digestData.d.GetContextWebInformation.FormDigestValue;

  // Step 3: Create list item
  const body = {
    __metadata: { type: entityType },
    Title: payload.Title,
    OrdererName: payload.OrdererName,
    OrdererEmail: payload.OrdererEmail,
    JobNameId: payload.JobNameId || null,
    Delivery_x0020_Location: payload.PickLocation,
    Delivery_x0020_Date: payload.TimeRequested || null,
    Quantities_x0020_Requested: payload.ItemsSummary,
    Additional_x0020_Notes_x002f_Com: payload.Notes || null,
    Status: "Received"
  };

  const postUrl =
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(ORDERS_LIST)}')/items`;

  const postRes = await fetch(postUrl, {
    method: "POST",
    headers: {
      "Accept": "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
      "Authorization": token,
      "X-RequestDigest": digest
    },
    body: JSON.stringify(body)
  });

  if (!postRes.ok) {
    const errData = await postRes.json().catch(() => ({}));
    throw new Error(errData?.error?.message?.value || `Submit failed: ${postRes.status}`);
  }

  return {
    statusCode: 200,
    headers,
    body: JSON.stringify({ success: true })
  };
}
