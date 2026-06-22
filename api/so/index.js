module.exports = async function (context, req) {
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHAREPOINT_SITE } = process.env;

  function err(status, msg) {
    context.res = { status, headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ error: msg }) };
  }
  function ok(data) {
    context.res = { status: 200,
                    headers: { 'Content-Type': 'application/json',
                               'Cache-Control': 'no-cache' },
                    body: JSON.stringify(data) };
  }

  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !SHAREPOINT_SITE) {
    return err(500, 'Missing server configuration. Check Azure app settings: TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHAREPOINT_SITE');
  }

  const action = (req.query.action || '').toLowerCase();
  if (!action) return err(400, 'Missing ?action= parameter. Use getassets or getbookings.');

  let token;
  try {
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type:    'client_credentials',
          client_id:     CLIENT_ID,
          client_secret: CLIENT_SECRET,
          scope:         'https://graph.microsoft.com/.default'
        }).toString()
      }
    );
    const tokenData = await tokenRes.json();
    if (!tokenRes.ok || !tokenData.access_token) {
      throw new Error(tokenData.error_description || tokenData.error || 'Token request failed');
    }
    token = tokenData.access_token;
  } catch (e) {
    return err(502, 'Could not authenticate with Microsoft: ' + e.message);
  }

  let siteId;
  try {
    const siteRes = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const siteData = await siteRes.json();
    if (!siteRes.ok) throw new Error(siteData.error?.message || 'Site lookup failed');
    siteId = siteData.id;
  } catch (e) {
    return err(502, 'Could not find SharePoint site: ' + e.message);
  }

  async function getListItems(listName) {
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}/items?expand=fields&$top=999`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const data = await res.json();
    if (!res.ok) throw new Error(data.error?.message || `${listName} list fetch failed`);
    return data.value || [];
  }

  if (action === 'getassets') {
    try {
      const rows = await getListItems('Assets');
      const assets = rows.map(item => {
        const f = item.fields;
        return {
          id:          f.AssetId         || '',
          name:        f.Title           || '',
          category:    f.Category        || 'Equipment',
          totalQty:    Number(f.TotalQty)    || 0,
          inRepairQty: Number(f.InRepairQty) || 0,
          location:    f.Location        || '',
          condition:   f.Condition       || 'Good',
          loanAllowed: f.LoanAllowed !== false,
          owningTeam:  f.OwningTeam      || '',
          specs:       f.Specs           || '',
          _spId:       item.id
        };
      });
      return ok(assets);
    } catch (e) {
      return err(502, 'Error reading Assets list: ' + e.message);
    }
  }

  if (action === 'getbookings') {
    try {
      const rows = await getListItems('Bookings');
      const bookings = rows.map(item => {
        const f = item.fields;
        return {
          id:          f.Ref             || ('REQ-' + item.id),
          assetId:     f.AssetId         || '',
          assetName:   f.AssetName       || '',
          qty:         Number(f.Qty)          || 0,
          returnedQty: f.ReturnedQty != null ? Number(f.ReturnedQty) : null,
          requester:   f.RequesterName   || '',
          email:       f.RequesterEmail  || '',
          event:       f.Event           || '',
          checkout:    (f.CheckoutDate   || '').split('T')[0],
          ret:         (f.ReturnDate     || '').split('T')[0],
          actual:      (f.ActualReturn   || '').split('T')[0],
          status:      f.Status          || 'Pending',
          note:        f.Note            || '',
          created:     (f.Created        || '').split('T')[0],
          _spId:       item.id
        };
      });
      return ok(bookings);
    } catch (e) {
      return err(502, 'Error reading Bookings list: ' + e.message);
    }
  }

  return err(400, `Unknown action "${action}". Valid actions: getassets, getbookings`);
};
