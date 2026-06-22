const { app } = require('@azure/functions');

app.http('sp', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHAREPOINT_SITE } = process.env;

        function errRes(status, msg) {
            return new Response(JSON.stringify({ error: msg }), {
                status, headers: { 'Content-Type': 'application/json' }
            });
        }
        function okRes(data) {
            return new Response(JSON.stringify(data), {
                status: 200,
                headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' }
            });
        }

        if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !SHAREPOINT_SITE) {
            return errRes(500, 'Missing server configuration.');
        }

        const action = (new URL(request.url).searchParams.get('action') || '').toLowerCase();
        if (!action) return errRes(400, 'Missing ?action= parameter.');

        // ── App-only token ───────────────────────────────────────────────
        let token;
        try {
            const tokenRes = await fetch(
                `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
                {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({
                        grant_type: 'client_credentials',
                        client_id: CLIENT_ID,
                        client_secret: CLIENT_SECRET,
                        scope: 'https://graph.microsoft.com/.default'
                    }).toString()
                }
            );
            const tokenData = await tokenRes.json();
            if (!tokenRes.ok || !tokenData.access_token) {
                throw new Error(tokenData.error_description || 'Token request failed');
            }
            token = tokenData.access_token;
        } catch (e) {
            return errRes(502, 'Auth failed: ' + e.message);
        }

        // ── SharePoint site ID ───────────────────────────────────────────
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
            return errRes(502, 'Site lookup failed: ' + e.message);
        }

        // ── Helpers ──────────────────────────────────────────────────────
        async function getListItems(listName) {
            const res = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}/items?expand=fields&$top=999`,
                { headers: { Authorization: `Bearer ${token}` } }
            );
            const data = await res.json();
            if (!res.ok) throw new Error(data.error?.message || `${listName} fetch failed`);
            return data.value || [];
        }

        function getUserFromHeader(req) {
            const principal = req.headers.get('x-ms-client-principal');
            if (!principal) return null;
            try {
                const decoded = Buffer.from(principal, 'base64').toString('utf-8');
                const parsed = JSON.parse(decoded);
                const email = parsed.userDetails || '';
                const nameClaim = (parsed.claims || []).find(c =>
                    c.typ === 'name' || c.typ === 'http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name'
                );
                const name = nameClaim?.val || email.split('@')[0] || 'User';
                return { email, name };
            } catch { return null; }
        }

        // SharePoint base URL from env var: wvuk.sharepoint.com:/sites/AssetTest
        const spBaseUrl = 'https://' + SHAREPOINT_SITE.replace(':/', '/');

        // ── ROUTE: getme ─────────────────────────────────────────────────
        if (action === 'getme') {
            const user = getUserFromHeader(request);
            if (!user) {
                return okRes({ name: 'User', email: '', isAdmin: false });
            }
            const adminEmails = (process.env.ADMIN_EMAILS || '')
                .split(',').map(e => e.trim().toLowerCase()).filter(Boolean);
            const isAdmin = adminEmails.includes(user.email.toLowerCase());
            return okRes({ name: user.name, email: user.email, isAdmin });
        }

        // ── ROUTE: getassets ─────────────────────────────────────────────
        if (action === 'getassets') {
            try {
                const rows = await getListItems('Assets');
                const assets = rows.map(item => {
                    const f = item.fields;
                    return {
                        id:          f.field_0         || '',
                        name:        f.Title           || '',
                        category:    f.field_2         || 'Equipment',
                        totalQty:    Number(f.field_3)     || 0,
                        inRepairQty: Number(f.field_4)     || 0,
                        location:    f.field_5         || '',
                        condition:   f.field_6         || 'Good',
                        loanAllowed: f.field_7 > 0,
                        owningTeam:  f.field_8         || '',
                        specs:       f.field_9         || '',
                        _spId:       item.id
                    };
                });
                return okRes(assets);
            } catch (e) {
                return errRes(502, 'Error reading Assets: ' + e.message);
            }
        }

        // ── ROUTE: getbookings ───────────────────────────────────────────
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
                return okRes(bookings);
            } catch (e) {
                return errRes(502, 'Error reading Bookings: ' + e.message);
            }
        }

        // ── ROUTE: createbooking (POST) ──────────────────────────────────
        if (action === 'createbooking') {
            try {
                const body = await request.json();

                // Generate next REQ number
                const existing = await getListItems('Bookings');
                let maxNum = 1000;
                existing.forEach(item => {
                    const ref = item.fields?.Ref || '';
                    const num = parseInt(ref.replace('REQ-', ''));
                    if (!isNaN(num) && num > maxNum) maxNum = num;
                });
                const ref = 'REQ-' + (maxNum + 1);

                const createRes = await fetch(
                    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/Bookings/items`,
                    {
                        method: 'POST',
                        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            fields: {
                                Title:          body.assetName || '',
                                Ref:            ref,
                                AssetId:        body.assetId   || '',
                                AssetName:      body.assetName || '',
                                Qty:            body.qty       || 1,
                                RequesterName:  body.requester || '',
                                RequesterEmail: body.email     || '',
                                Event:          body.event     || '',
                                CheckoutDate:   body.checkout  ? body.checkout + 'T00:00:00Z' : null,
                                ReturnDate:     body.ret       ? body.ret + 'T00:00:00Z'      : null,
                                Status:         'Pending',
                                Note:           body.note      || ''
                            }
                        })
                    }
                );
                const created = await createRes.json();
                if (!createRes.ok) throw new Error(created.error?.message || 'Create failed');
                return okRes({ id: ref, _spId: created.id });
            } catch (e) {
                return errRes(502, 'Error creating booking: ' + e.message);
            }
        }

        // ── ROUTE: updatebooking (POST) ──────────────────────────────────
        if (action === 'updatebooking') {
            try {
                const body = await request.json();
                const { spId, status, returnedQty, actualReturn } = body;
                if (!spId) return errRes(400, 'Missing spId');

                const fields = { Status: status };
                if (returnedQty !== undefined && returnedQty !== null) fields.ReturnedQty = returnedQty;
                if (actualReturn) fields.ActualReturn = actualReturn + 'T00:00:00Z';

                const patchRes = await fetch(
                    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/Bookings/items/${spId}`,
                    {
                        method: 'PATCH',
                        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
                        body: JSON.stringify({ fields })
                    }
                );
                if (!patchRes.ok) {
                    const err = await patchRes.json().catch(() => ({}));
                    throw new Error(err.error?.message || 'Update failed: HTTP ' + patchRes.status);
                }
                return okRes({ success: true });
            } catch (e) {
                return errRes(502, 'Error updating booking: ' + e.message);
            }
        }

        return errRes(400, `Unknown action "${action}".`);
    }
});
