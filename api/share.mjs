// Short share links backed by Vercel Blob. Requires a Blob store connected
// to the project (Storage tab) which injects BLOB_READ_WRITE_TOKEN.
import { put, del, list } from '@vercel/blob';
const cleanSlug = s => String(s || '').toLowerCase().replace(/[^a-z0-9-]/g, '').slice(0, 60);
export default async function handler(req, res) {
  if (!process.env.BLOB_READ_WRITE_TOKEN) return res.status(404).json({ error: 'storage not configured' });
  const slug = cleanSlug((req.query && req.query.slug) || (req.body && req.body.slug));
  if (!slug) return res.status(400).json({ error: 'bad slug' });
  try {
    if (req.method === 'POST') {
      const { payload } = req.body || {};
      if (!payload || JSON.stringify(payload).length > 900000) return res.status(400).json({ error: 'bad payload' });
      const doc = { exp: Date.now() + 14 * 864e5, on: true, payload };
      await put(`p/${slug}.json`, JSON.stringify(doc), { access: 'public', addRandomSuffix: false, allowOverwrite: true, contentType: 'application/json', cacheControlMaxAge: 60 });
      const origin = (req.headers['x-forwarded-proto'] || 'https') + '://' + req.headers.host;
      return res.status(200).json({ url: origin + '/p/' + slug, exp: doc.exp, slug });
    }
    if (req.method === 'DELETE') {
      const { blobs } = await list({ prefix: `p/${slug}.json`, limit: 1 });
      if (blobs.length) await del(blobs[0].url);
      return res.status(200).json({ ok: true });
    }
    return res.status(405).json({ error: 'POST or DELETE' });
  } catch (e) {
    return res.status(500).json({ error: String(e.message || e) });
  }
}
