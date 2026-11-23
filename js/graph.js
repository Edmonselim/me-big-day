// Microsoft sign-in (OAuth 2.0 auth-code + PKCE, no library) and OneDrive file access via Microsoft Graph.
import { CONFIG } from '../config.js';

const AUTH = `https://login.microsoftonline.com/${CONFIG.authority}/oauth2/v2.0`;
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = 'openid profile offline_access User.Read Files.ReadWrite.All';
const TOKENS = 'meb.tokens';

const redirectUri = () => location.origin + location.pathname.replace(/index\.html$/, '');

function b64url(bytes) {
  return btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

export async function signIn(prompt) {
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = b64url(crypto.getRandomValues(new Uint8Array(12)));
  sessionStorage.setItem('meb.pkce', JSON.stringify({ verifier, state }));
  localStorage.setItem('meb.pkce', JSON.stringify({ verifier, state })); // iOS home-screen apps may drop sessionStorage on redirect
  const p = new URLSearchParams({
    client_id: CONFIG.clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query',
    scope: SCOPES, code_challenge: challenge, code_challenge_method: 'S256', state,
  });
  if (prompt) p.set('prompt', prompt);
  location.assign(`${AUTH}/authorize?${p}`);
}

export function signOut() {
  localStorage.removeItem(TOKENS);
  location.assign(`${AUTH}/logout?post_logout_redirect_uri=${encodeURIComponent(redirectUri())}`);
}

// Call on page load: finishes a sign-in redirect if one is in progress.
export async function handleRedirect() {
  const q = new URLSearchParams(location.search);
  if (q.get('error')) {
    history.replaceState(null, '', redirectUri());
    throw new Error(q.get('error_description') || q.get('error'));
  }
  if (!q.get('code')) return false;
  const saved = JSON.parse(sessionStorage.getItem('meb.pkce') || localStorage.getItem('meb.pkce') || 'null');
  history.replaceState(null, '', redirectUri());
  if (!saved || saved.state !== q.get('state')) throw new Error('Sign-in expired, please try again.');
  await tokenRequest({ grant_type: 'authorization_code', code: q.get('code'), code_verifier: saved.verifier, redirect_uri: redirectUri() });
  localStorage.removeItem('meb.pkce');
  sessionStorage.removeItem('meb.pkce');
  return true;
}

async function tokenRequest(params) {
  const res = await fetch(`${AUTH}/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: CONFIG.clientId, scope: SCOPES, ...params }),
  });
  const j = await res.json();
  if (!res.ok) { const e = new Error(j.error_description || j.error || 'Sign-in failed'); e.code = j.error; throw e; }
  const t = { access: j.access_token, refresh: j.refresh_token, exp: Date.now() + (j.expires_in - 120) * 1000 };
  localStorage.setItem(TOKENS, JSON.stringify(t));
  return t;
}

export const isSignedIn = () => !!localStorage.getItem(TOKENS);

export class AuthRequired extends Error {}

async function accessToken() {
  const t = JSON.parse(localStorage.getItem(TOKENS) || 'null');
  if (!t) throw new AuthRequired('Not signed in');
  if (Date.now() < t.exp) return t.access;
  try {
    return (await tokenRequest({ grant_type: 'refresh_token', refresh_token: t.refresh })).access;
  } catch (e) {
    if (e.code === 'invalid_grant' || e.code === 'interaction_required') { localStorage.removeItem(TOKENS); throw new AuthRequired('Session expired'); }
    throw e;
  }
}

export class GraphError extends Error {
  constructor(status, msg, code) { super(msg); this.status = status; this.code = code; }
}

async function graph(path, opts = {}) {
  const token = await accessToken();
  const res = await fetch(path.startsWith('http') ? path : GRAPH + path, {
    ...opts, headers: { Authorization: `Bearer ${token}`, ...(opts.headers || {}) },
  });
  if (res.status === 401) { localStorage.removeItem(TOKENS); throw new AuthRequired('Session expired'); }
  if (!res.ok) {
    let msg = res.statusText, code;
    try { const j = await res.json(); msg = j.error?.message || msg; code = j.error?.code; } catch { /* ignore */ }
    throw new GraphError(res.status, msg, code);
  }
  return res;
}

export async function me() { return (await graph('/me')).json(); }

const encPath = (p) => p.split('/').map(encodeURIComponent).join('/');
const SELECT = '?select=id,name,eTag,cTag,size,lastModifiedDateTime,lastModifiedBy,parentReference,webUrl,@microsoft.graph.downloadUrl';

// Owner: find the file by its OneDrive path.
export async function findByPath(path) {
  const it = await (await graph(`/me/drive/root:/${encPath(path)}${SELECT}`)).json();
  return { driveId: it.parentReference.driveId, itemId: it.id, name: it.name };
}

// Partner: a share link, a partner code (driveId!itemId) or "Shared with me".
export async function findShared({ link, code, name }) {
  if (code) {
    // partner code format: driveId|itemId (shown on the owner's Settings screen)
    const [driveId, itemId] = code.trim().split('|');
    if (!driveId || !itemId) throw new Error('That partner code looks incomplete – copy it again from the Settings screen.');
    const it = await (await graph(`/drives/${driveId}/items/${encodeURIComponent(itemId)}${SELECT}`)).json();
    return { driveId: it.parentReference.driveId, itemId: it.id, name: it.name };
  }
  if (link) {
    const id = 'u!' + btoa(unescape(encodeURIComponent(link.trim()))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
    const it = await (await graph(`/shares/${id}/driveItem`, { headers: { Prefer: 'redeemSharingLink' } })).json();
    return { driveId: it.parentReference.driveId, itemId: it.id, name: it.name };
  }
  const list = await (await graph('/me/drive/sharedWithMe')).json();
  const hit = (list.value || []).find((x) => (x.remoteItem?.name || x.name) === name);
  if (!hit) throw new Error(`"${name}" isn't in your Shared with me list yet. Ask for the share link or partner code instead.`);
  const r = hit.remoteItem || hit;
  return { driveId: r.parentReference.driveId, itemId: r.id, name: r.name };
}

export async function itemMeta(file) {
  return (await graph(`/drives/${file.driveId}/items/${file.itemId}${SELECT}`)).json();
}

export async function download(file) {
  const meta = await itemMeta(file);
  const url = meta['@microsoft.graph.downloadUrl'];
  const res = url ? await fetch(url, { cache: 'no-store' }) : await graph(`/drives/${file.driveId}/items/${file.itemId}/content`);
  if (!res.ok) throw new GraphError(res.status, 'Download failed');
  return { bytes: new Uint8Array(await res.arrayBuffer()), meta };
}

// Upload new contents; If-Match makes sure nobody saved in between.
export async function upload(file, bytes, eTag) {
  const type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  if (bytes.length < 3.9 * 1024 * 1024) {
    const res = await graph(`/drives/${file.driveId}/items/${file.itemId}/content`, {
      method: 'PUT', body: bytes, headers: { 'Content-Type': type, ...(eTag ? { 'If-Match': eTag } : {}) },
    });
    return res.json();
  }
  const sess = await (await graph(`/drives/${file.driveId}/items/${file.itemId}/createUploadSession`, {
    method: 'POST', headers: { 'Content-Type': 'application/json', ...(eTag ? { 'If-Match': eTag } : {}) },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
  })).json();
  const CHUNK = 320 * 1024 * 10; // multiple of 320 KiB
  let last;
  for (let start = 0; start < bytes.length; start += CHUNK) {
    const end = Math.min(start + CHUNK, bytes.length);
    last = await fetch(sess.uploadUrl, {
      method: 'PUT', body: bytes.subarray(start, end),
      headers: { 'Content-Range': `bytes ${start}-${end - 1}/${bytes.length}` },
    });
    if (!last.ok && last.status !== 202) throw new GraphError(last.status, 'Upload failed');
  }
  return last.json();
}
