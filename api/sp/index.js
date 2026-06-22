const { app } = require('@azure/functions');

app.http('sp', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHAREPOINT_SITE } = process.env;

        function errRes(status, msg) {
            return new Response(JSON.stringify({ error: msg }), {
                status,
                headers: { 'Content-Type': 'application/json' }
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

        async function getListItems(listName) {
            const res = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}/items?expand=fields&$top=999`,
                { headers: { Authorization: `Bearer ${token}` } }
            );
            const data = await res.json();
            if (!res.ok) throw new Error(data.error?.message || `${listName} fetch failed`);
            return data.value || [];
        }

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

        return errRes(400, `Unknown action "${action}".`);
    }
});
