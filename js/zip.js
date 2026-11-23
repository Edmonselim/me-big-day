// Minimal ZIP reader/writer for .xlsx files.
// Uses the platform's CompressionStream / DecompressionStream (Safari 16.4+, Node 18+),
// so the app needs no third-party libraries. Unchanged entries are copied byte-for-byte.

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

async function streamThrough(bytes, stream) {
  const blob = new Blob([bytes]);
  const out = await new Response(blob.stream().pipeThrough(stream)).arrayBuffer();
  return new Uint8Array(out);
}
const inflate = (b) => streamThrough(b, new DecompressionStream('deflate-raw'));
const deflate = (b) => streamThrough(b, new CompressionStream('deflate-raw'));

const td = new TextDecoder();
const te = new TextEncoder();

export class Zip {
  constructor() {
    this.entries = new Map(); // name -> entry
    this.order = [];
  }

  static async read(buffer) {
    const u8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Not a valid .xlsx file (zip directory not found).');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    if (p === 0xffffffff) throw new Error('ZIP64 workbooks are not supported.');
    const zip = new Zip();
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Corrupt zip directory.');
      const method = dv.getUint16(p + 10, true);
      const time = dv.getUint16(p + 12, true);
      const date = dv.getUint16(p + 14, true);
      const crc = dv.getUint32(p + 16, true);
      const csize = dv.getUint32(p + 20, true);
      const usize = dv.getUint32(p + 24, true);
      const nlen = dv.getUint16(p + 28, true);
      const xlen = dv.getUint16(p + 30, true);
      const clen = dv.getUint16(p + 32, true);
      const flags = dv.getUint16(p + 8, true);
      const lho = dv.getUint32(p + 42, true);
      const name = td.decode(u8.subarray(p + 46, p + 46 + nlen));
      p += 46 + nlen + xlen + clen;
      const lnlen = dv.getUint16(lho + 26, true);
      const lxlen = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnlen + lxlen;
      zip.entries.set(name, {
        name, method, time, date, crc, csize, usize, flags,
        raw: u8.subarray(start, start + csize), data: null, changed: false,
      });
      zip.order.push(name);
    }
    return zip;
  }

  has(name) { return this.entries.has(name); }
  names() { return [...this.order]; }

  async bytes(name) {
    const e = this.entries.get(name);
    if (!e) return null;
    if (!e.data) e.data = e.method === 0 ? e.raw : await inflate(e.raw);
    return e.data;
  }

  async text(name) {
    const b = await this.bytes(name);
    return b ? td.decode(b) : null;
  }

  setText(name, text) { this.setBytes(name, te.encode(text)); }

  setBytes(name, bytes) {
    const e = this.entries.get(name);
    if (e) { e.data = bytes; e.changed = true; }
    else {
      this.entries.set(name, { name, method: 8, time: 0, date: 0x21, flags: 0, data: bytes, changed: true });
      this.order.push(name);
    }
  }

  remove(name) {
    this.entries.delete(name);
    this.order = this.order.filter((n) => n !== name);
  }

  async write() {
    const parts = [];
    const central = [];
    let offset = 0;
    for (const name of this.order) {
      const e = this.entries.get(name);
      if (e.changed) {
        e.crc = crc32(e.data);
        e.usize = e.data.length;
        e.raw = await deflate(e.data);
        e.method = 8;
        e.csize = e.raw.length;
        e.flags = 0;
        e.changed = false;
      }
      const nameBytes = te.encode(name);
      const flags = e.flags & ~0x0008 | (/[^\x00-\x7f]/.test(name) ? 0x0800 : 0);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true);
      lh.setUint16(6, flags, true);
      lh.setUint16(8, e.method, true);
      lh.setUint16(10, e.time, true);
      lh.setUint16(12, e.date, true);
      lh.setUint32(14, e.crc, true);
      lh.setUint32(18, e.csize, true);
      lh.setUint32(22, e.usize, true);
      lh.setUint16(26, nameBytes.length, true);
      lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), nameBytes, e.raw);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true);
      ch.setUint16(4, 20, true);
      ch.setUint16(6, 20, true);
      ch.setUint16(8, flags, true);
      ch.setUint16(10, e.method, true);
      ch.setUint16(12, e.time, true);
      ch.setUint16(14, e.date, true);
      ch.setUint32(16, e.crc, true);
      ch.setUint32(20, e.csize, true);
      ch.setUint32(24, e.usize, true);
      ch.setUint16(28, nameBytes.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), nameBytes);
      offset += 30 + nameBytes.length + e.raw.length;
    }
    const cdSize = central.reduce((s, b) => s + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, this.order.length, true);
    end.setUint16(10, this.order.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(end.buffer)];
    const total = all.reduce((s, b) => s + b.length, 0);
    const out = new Uint8Array(total);
    let o = 0;
    for (const b of all) { out.set(b, o); o += b.length; }
    return out;
  }
}
