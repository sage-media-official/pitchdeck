import { list } from '@vercel/blob';
export default async function handler(req, res) {
  if (!process.env.BLOB_READ_WRITE_TOKEN) return res.status(404).json({ error: 'storage not configured' });
  const slug = String(req.query.slug || '').toLowerCase().replace(/[^a-z0-9-]/g, '').slice(0, 60);
  if (!slug) return res.status(400).json({ error: 'bad slug' });
  try {
    const { blobs } = await list({ prefix: `p/${slug}.json`, limit: 1 });
    if (!blobs.length) return res.status(410).json({ error: 'gone' });
    const doc = await (await fetch(blobs[0].url + '?t=' + Date.now())).json();
    if (doc.on === false || Date.now() > doc.exp) return res.status(410).json({ error: 'expired' });
    res.setHeader('cache-control', 'no-store');
    return res.status(200).json(doc);
  } catch (e) {
    return res.status(500).json({ error: String(e.message || e) });
  }
}
