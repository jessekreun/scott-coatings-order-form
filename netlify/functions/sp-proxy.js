// sp-proxy.js — SharePoint proxy for Scott Coatings Order Form
// Handles all SP reads and writes server-side to avoid CORS restrictions.

const SITE = "https://scottcoatings.sharepoint.com/sites/ProjectOperations";
const MPI_LIST = "Master Project Index";
const ORDERS_LIST = "Project Orders";
const PAINT_ORDERS_LIST = "Paint Orders";
const ACTIVE_STATUSES = ["Course of Construction", "Pre-Construction"];

exports.handler = async function(event) {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Authorization, Content-Type",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Content-Type": "application/json"
  };

  if (event.httpMethod === "OPTIONS") return { statusCode: 200, headers, body: "" };

  const action = event.queryStringParameters?.action;
  const token = event.headers?.authorization;

  if (!token) return { statusCode: 401, headers, body: JSON.stringify({ error: "No authorization token provided" }) };
  if (!action) return { statusCode: 400, headers, body: JSON.stringify({ error: "No action specified" }) };

  try {
    switch (action) {
      case "getMPI":          return await getMPI(token, headers);
      case "getItems":        return await getItems(token, headers);
      case "getMAT":          return await getMAT(token, headers);
      case "submitOrder":     return await submitOrder(token, headers, event.body);
      case "submitPaintOrder": return await submitPaintOrder(token, headers, event.body);
      default:
        return { statusCode: 400, headers, body: JSON.stringify({ error: `Unknown action: ${action}` }) };
    }
  } catch (err) {
    console.error(`Error in action ${action}:`, err);
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message || "Internal server error" }) };
  }
};

// ── SP HELPERS ────────────────────────────────────────────────────────────────

async function spGetAll(baseUrl, token) {
  let results = [];
  let url = baseUrl;
  while (url) {
    const res = await fetch(url, {
      headers: { "Accept": "application/json;odata=verbose", "Authorization": token }
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

async function getEntityType(listName, token) {
  const res = await fetch(
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')` +
    `?$select=ListItemEntityTypeFullName`,
    { headers: { "Accept": "application/json;odata=verbose", "Authorization": token } }
  );
  if (!res.ok) throw new Error(`Meta fetch failed: ${res.status}`);
  const data = await res.json();
  return data.d.ListItemEntityTypeFullName;
}

async function getDigest(token) {
  const res = await fetch(`${SITE}/_api/contextinfo`, {
    method: "POST",
    headers: { "Accept": "application/json;odata=verbose", "Authorization": token }
  });
  if (!res.ok) throw new Error(`Digest fetch failed: ${res.status}`);
  const data = await res.json();
  return data.d.GetContextWebInformation.FormDigestValue;
}

async function spPostItem(listName, body, token) {
  const [entityType, digest] = await Promise.all([
    getEntityType(listName, token),
    getDigest(token)
  ]);
  const res = await fetch(
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items`,
    {
      method: "POST",
      headers: {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "Authorization": token,
        "X-RequestDigest": digest
      },
      body: JSON.stringify({ __metadata: { type: entityType }, ...body })
    }
  );
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message?.value || `Submit failed: ${res.status}`);
  }
  return res.json();
}

// ── GET MPI ───────────────────────────────────────────────────────────────────

async function getMPI(token, headers) {
  const statusFilter = ACTIVE_STATUSES
    .map(s => `Project_x0020_Status eq '${encodeURIComponent(s)}'`)
    .join(" or ");

  const results = await spGetAll(
    `${SITE}/_api/web/lists/getbytitle('${encodeURIComponent(MPI_LIST)}')/items` +
    `?$select=Id,Title,Project_x0020_Name,Project_x0020_Location,PM_x0020_Name,Client_x0020_Company/Title,Project_x0020_Status` +
    `&$expand=Client_x0020_Company` +
    `&$filter=${statusFilter}` +
    `&$orderby=Project_x0020_Name` +
    `&$top=500`,
    token
  );

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

  return { statusCode: 200, headers, body: JSON.stringify({ jobs }) };
}

// ── GET ITEMS (SUN/PPE list) ──────────────────────────────────────────────────

async function getItems(token, headers) {
  const results = await spGetAll(
    `${SITE}/_api/web/GetList(@listUrl)/items` +
    `?@listUrl='${encodeURIComponent("/sites/ProjectOperations/Lists/SUN")}'` +
    `&$select=Id,Title,Description,SageID,Manufacturer,VendorID,Packaging,ImageURL` +
    `&$top=500`,
    token
  );

  const items = results.map(r => ({
    id: r.Id,
    sageId: r.Title || "",
    name: r.Description || r.Title || "Unnamed item",
    type: r.SageID || "",
    mfr: r.Manufacturer || "",
    vendorId: r.VendorID || "",
    pkg: r.Packaging || "",
    picture: r.ImageURL || null
  }));

  return { statusCode: 200, headers, body: JSON.stringify({ items }) };
}

// ── GET MAT (paint/materials list) ───────────────────────────────────────────

async function getMAT(token, headers) {
  const results = await spGetAll(
    `${SITE}/_api/web/GetList(@listUrl)/items` +
    `?@listUrl='${encodeURIComponent("/sites/ProjectOperations/Lists/2MAT")}'` +
    `&$select=Id,Title,Description,Manufacturer,VendorID,SageID,Paint_x002f_Coatings_x0020_Type,Packaging,Cost_x0020__x002f__x0020_Gal` +
    `&$top=500`,
    token
  );

  const items = results.map(r => ({
    id: r.Id,
    sageId: r.SageID || r.Title || "",
    name: r.Description || r.Title || "Unnamed product",
    mfr: r.Manufacturer || "",
    vendorId: r.VendorID || "",
    type: r.Paint_x002f_Coatings_x0020_Type || "",
    pkg: r.Packaging || "",
    costPerGal: r.Cost_x0020__x002f__x0020_Gal || null
  }));

  return { statusCode: 200, headers, body: JSON.stringify({ items }) };
}

// ── SUBMIT SUNDRY/PPE ORDER ───────────────────────────────────────────────────

async function submitOrder(token, headers, rawBody) {
  if (!rawBody) throw new Error("No body provided");
  const p = JSON.parse(rawBody);

  await spPostItem(ORDERS_LIST, {
    Title: p.Title,
    OrdererName: p.OrdererName,
    OrdererEmail: p.OrdererEmail,
    JobNameId: p.JobNameId || null,
    Delivery_x0020_Location: p.PickLocation,
    Delivery_x0020_Date: p.TimeRequested || null,
    Quantities_x0020_Requested: p.SunOeqSummary || null,
    PPEItemsOrdered: p.PPESummary || null,
    Additional_x0020_Notes_x002f_Com: p.Notes || null,
    Status: "Received"
  }, token);

  return { statusCode: 200, headers, body: JSON.stringify({ success: true }) };
}

// ── SUBMIT PAINT ORDER ────────────────────────────────────────────────────────

async function submitPaintOrder(token, headers, rawBody) {
  if (!rawBody) throw new Error("No body provided");
  const p = JSON.parse(rawBody);

  await spPostItem(PAINT_ORDERS_LIST, {
    Title: p.Title,
    OrdererName: p.OrdererName,
    OrdererEmail: p.OrdererEmail,
    ProjectNameId: p.JobNameId || null,
    PickuporDelivery: p.PickupOrDelivery,
    DeliveryDate: p.DeliveryDate || null,
    DeliveryNotes: p.DeliveryNotes || null,
    ItemsOrdered: p.ItemsSummary || null,
    ItemCount: p.ItemCount || 0,
    Urgent: p.Urgent || false,
    Status: "Received"
  }, token);

  return { statusCode: 200, headers, body: JSON.stringify({ success: true }) };
}
