/* impactTV 端末管理 — デモモード（demo-source.html から抽出・自動生成）
 * 構成: GasMock → __makeServer → __LEDGER__ → window.__startDemo__
 * 読み込み後に window.__startDemo__() を呼ぶと、メモリ上のモックGAS環境に
 * 台帳を取り込み、window.__MOCK_API__ を設置します（合言葉: kaino）。
 * 記録はメモリ上だけに保持され、再読み込みで初期状態に戻ります。
 */

/* In-memory mock of the Google Apps Script services used by Code.gs / Admin.gs.
 * Works in Node and in the browser. Not a full emulation – just enough fidelity
 * for the calls the app makes (values, formulas, TextFinder, rows, cache, lock…). */
(function (root) {
  'use strict';

  function createGasEnv(opts) {
    opts = opts || {};
    const log = { alerts: [], prompts: [], dialogs: [], trashed: [], lockCalls: 0 };
    let uuidSeq = 0;
    const promptAnswers = (opts.promptAnswers || []).slice();

    // ---------------- helpers ----------------
    function colToNum(letters) {
      let n = 0;
      for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
      return n;
    }
    function parseA1(a1, sheet) {
      const m = /^(?:'?([^'!]+)'?!)?([A-Z]*)(\d*)(?::([A-Z]*)(\d*))?$/.exec(a1.replace(/\$/g, ''));
      if (!m) throw new Error('bad A1: ' + a1);
      const c1 = m[2] ? colToNum(m[2]) : 1;
      const r1 = m[3] ? Number(m[3]) : 1;
      let c2, r2;
      if (m[4] !== undefined || m[5] !== undefined) {
        c2 = m[4] ? colToNum(m[4]) : sheet.maxCols;
        r2 = m[5] ? Number(m[5]) : sheet.maxRows;
      } else { c2 = c1; r2 = r1; }
      if (!m[3] && m[2]) { /* column-only like "A:A" */ }
      return { row: r1, col: c1, numRows: r2 - r1 + 1, numCols: c2 - c1 + 1 };
    }
    function cloneVal(v) { return v instanceof Date ? new Date(v.getTime()) : v; }
    function blank(cell) { return !cell || ((cell.v === '' || cell.v == null) && !cell.f); }

    // ---------------- TextFinder ----------------
    class TextFinder {
      constructor(range, text) { this.range = range; this.text = String(text); this.entire = false; this.cs = false; }
      matchEntireCell(b) { this.entire = !!b; return this; }
      matchCase(b) { this.cs = !!b; return this; }
      findAll() {
        const out = [];
        const r = this.range;
        for (let i = 0; i < r.numRows; i++) {
          for (let j = 0; j < r.numCols; j++) {
            const cell = r.sheet._cell(r.row + i, r.col + j, false);
            if (!cell) continue;
            let val = cell.v instanceof Date ? cell.v.toString() : String(cell.v == null ? '' : cell.v);
            let q = this.text;
            if (!this.cs) { val = val.toLowerCase(); q = q.toLowerCase(); }
            const hit = this.entire ? val === q : (q !== '' && val.indexOf(q) >= 0);
            if (hit) out.push(r.sheet.getRange(r.row + i, r.col + j, 1, 1));
          }
        }
        return out;
      }
    }

    // ---------------- Range ----------------
    function makeRange(sheet, row, col, numRows, numCols) {
      const r = {
        sheet, row, col, numRows, numCols,
        getRow() { return row; },
        getColumn() { return col; },
        getNumRows() { return numRows; },
        getNumColumns() { return numCols; },
        getSheet() { return sheet; },
        getA1Notation() { return `R${row}C${col}:${numRows}x${numCols}`; },
        getValues() {
          const out = [];
          for (let i = 0; i < numRows; i++) {
            const line = [];
            for (let j = 0; j < numCols; j++) {
              const c = sheet._cell(row + i, col + j, false);
              line.push(c ? cloneVal(c.v == null ? '' : c.v) : '');
            }
            out.push(line);
          }
          return out;
        },
        getValue() { return this.getValues()[0][0]; },
        getFormulas() {
          const out = [];
          for (let i = 0; i < numRows; i++) {
            const line = [];
            for (let j = 0; j < numCols; j++) {
              const c = sheet._cell(row + i, col + j, false);
              line.push(c && c.f ? c.f : '');
            }
            out.push(line);
          }
          return out;
        },
        setValues(values) {
          if (!Array.isArray(values) || values.length !== numRows) throw new Error(`setValues: rows ${values && values.length} != ${numRows}`);
          values.forEach((line, i) => {
            if (!Array.isArray(line) || line.length !== numCols) throw new Error(`setValues: cols ${line && line.length} != ${numCols} (row ${i})`);
            line.forEach((v, j) => sheet._write(row + i, col + j, v));
          });
          return proxy;
        },
        setValue(v) {
          for (let i = 0; i < numRows; i++) for (let j = 0; j < numCols; j++) sheet._write(row + i, col + j, v);
          return proxy;
        },
        setFormula(f) { sheet._write(row, col, f); return proxy; },
        setNumberFormat(fmt) {
          for (let i = 0; i < Math.min(numRows, 5000); i++) for (let j = 0; j < numCols; j++) sheet._fmt[(row + i) + ':' + (col + j)] = fmt;
          return proxy;
        },
        clearContent() {
          for (let i = 0; i < numRows; i++) for (let j = 0; j < numCols; j++) sheet._write(row + i, col + j, '');
          return proxy;
        },
        createTextFinder(text) { return new TextFinder(this, text); },
        setDataValidation(rule) { sheet._validations.push({ row, col, numRows, numCols, rule }); return proxy; },
        setNote(n) { sheet._notes[row + ':' + col] = n; return proxy; },
      };
      const proxy = new Proxy(r, {
        get(t, prop) {
          if (prop in t) return t[prop];
          if (typeof prop === 'symbol') return undefined;
          return function () { return proxy; }; // formatting no-ops
        },
      });
      return proxy;
    }

    // ---------------- Sheet ----------------
    function makeSheet(ss, name) {
      const sh = {
        _rows: [], _fmt: {}, _validations: [], _notes: {}, _cf: [], _hidden: [],
        name, maxRows: 1000, maxCols: 26, frozenRows: 0, frozenCols: 0,
        _cell(r, c, create) {
          if (!this._rows[r - 1]) { if (!create) return null; this._rows[r - 1] = []; }
          const line = this._rows[r - 1];
          if (!line[c - 1]) { if (!create) return null; line[c - 1] = { v: '' }; }
          return line[c - 1];
        },
        _write(r, c, v) {
          if (r > this.maxRows) this.maxRows = r;
          if (c > this.maxCols) this.maxCols = c;
          const cell = this._cell(r, c, true);
          cell.f = '';
          if (typeof v === 'string' && v.charAt(0) === '=') { cell.f = v; cell.v = ''; return; }
          if (typeof v === 'string' && v.charAt(0) === "'") { cell.v = v.slice(1); return; }
          const isText = this._fmt[r + ':' + c] === '@';
          if (typeof v === 'string' && !isText && /^-?\d+(\.\d+)?$/.test(v.trim()) && v.trim() !== '') { cell.v = Number(v); return; }
          if (typeof v === 'boolean' || typeof v === 'number' || v instanceof Date) { cell.v = cloneVal(v); return; }
          if (v == null) { cell.v = ''; return; }
          if (typeof v === 'object') throw new Error('Cannot write object to cell: ' + JSON.stringify(v).slice(0, 80));
          cell.v = String(v);
        },
        getName() { return this.name; },
        setName(n) { this.name = n; return this; },
        getParent() { return ss; },
        getRange(a, b, c, d) {
          if (typeof a === 'string') {
            const p = parseA1(a, this);
            return makeRange(this, p.row, p.col, p.numRows, p.numCols);
          }
          if (!(a >= 1) || !(b >= 1)) throw new Error(`getRange: invalid start ${a},${b}`);
          const nr = c == null ? 1 : c, nc = d == null ? 1 : d;
          if (!(nr >= 1) || !(nc >= 1)) throw new Error(`getRange: invalid size ${nr}x${nc}`);
          return makeRange(this, a, b, nr, nc);
        },
        getLastRow() {
          for (let i = this._rows.length; i >= 1; i--) {
            const line = this._rows[i - 1];
            if (line && line.some((c) => !blank(c))) return i;
          }
          return 0;
        },
        getLastColumn() {
          let max = 0;
          this._rows.forEach((line) => {
            if (!line) return;
            for (let j = line.length; j >= 1; j--) if (!blank(line[j - 1])) { max = Math.max(max, j); break; }
          });
          return max;
        },
        getMaxRows() { return this.maxRows; },
        getMaxColumns() { return this.maxCols; },
        getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); },
        deleteRow(r) { this._rows.splice(r - 1, 1); this.maxRows--; return this; },
        insertRowsAfter(after, n) { this.maxRows += n; return this; },
        clear() { this._rows = []; this._fmt = {}; this._cf = []; return this; },
        clearContents() { this._rows = []; return this; },
        setConditionalFormatRules(rules) { this._cf = rules; return this; },
        getConditionalFormatRules() { return this._cf; },
        setFrozenRows(n) { this.frozenRows = n; return this; },
        setFrozenColumns(n) { this.frozenCols = n; return this; },
        hideColumns(c) { this._hidden.push(c); return this; },
        // helpers for tests
        _dump() { return this.getDataRange().getValues(); },
      };
      return new Proxy(sh, {
        get(t, prop) {
          if (prop in t) return t[prop];
          if (typeof prop === 'symbol') return undefined;
          return function () { return t; };
        },
      });
    }

    // ---------------- Spreadsheet ----------------
    const spreadsheets = {};
    function makeSpreadsheet(id, name) {
      const ss = {
        _id: id, _name: name, sheets: [], active: null, named: {}, tz: 'Etc/GMT',
        getId() { return this._id; },
        getName() { return this._name; },
        getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this._id + '/edit'; },
        getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; },
        getSheets() { return this.sheets.slice(); },
        insertSheet(n, index) {
          if (this.getSheetByName(n)) throw new Error('sheet exists: ' + n);
          const s = makeSheet(this, n);
          if (index == null || index > this.sheets.length) this.sheets.push(s);
          else this.sheets.splice(index, 0, s);
          return s;
        },
        deleteSheet(s) { this.sheets = this.sheets.filter((x) => x !== s); },
        setActiveSheet(s) { this.active = s; return s; },
        getActiveSheet() { return this.active || this.sheets[0]; },
        moveActiveSheet(pos) {
          const s = this.active;
          this.sheets = this.sheets.filter((x) => x !== s);
          this.sheets.splice(pos - 1, 0, s);
        },
        setSpreadsheetTimeZone(tz) { this.tz = tz; },
        getSpreadsheetTimeZone() { return this.tz; },
        setNamedRange(n, r) { if (this.named[n]) throw new Error('named range exists: ' + n); this.named[n] = r; },
        getNamedRanges() { const self = this; return Object.keys(this.named).map((n) => ({ getName: () => n, setRange: (r) => { self.named[n] = r; }, getRange: () => self.named[n] })); },
        getRangeByName(n) { return this.named[n] || null; },
      };
      spreadsheets[id] = ss;
      return ss;
    }

    const activeSS = makeSpreadsheet(opts.ssId || 'SS_TEST_1', opts.ssName || 'impactTV端末管理');
    activeSS.insertSheet(opts.defaultSheetName || 'シート1');

    // ---------------- Drive ----------------
    const files = {};
    const folders = {};
    let fileSeq = 0;
    function makeFolder(id, name) {
      const f = {
        _id: id, _name: name, files: [],
        getId() { return this._id; },
        getName() { return this._name; },
        createFolder(n) { return makeFolder('FOLDER_' + (++fileSeq), n); },
        createFile(blob) {
          const id = 'FILE_' + (++fileSeq);
          const file = {
            _id: id, _name: blob.getName(), blob, trashed: false, desc: '',
            getId() { return id; },
            getName() { return this._name; },
            getUrl() { return 'https://drive.google.com/file/d/' + id + '/view'; },
            setDescription(d) { this.desc = d; return this; },
            setTrashed(b) { this.trashed = b; log.trashed.push(id); return this; },
            getParents() { return iter([f]); },
          };
          files[id] = file;
          this.files.push(file);
          return file;
        },
      };
      folders[id] = f;
      return f;
    }
    const rootFolder = makeFolder('ROOT', 'マイドライブ');
    function iter(list) { let i = 0; return { hasNext: () => i < list.length, next: () => list[i++] }; }
    files[activeSS.getId()] = { getParents: () => iter([rootFolder]), getId: () => activeSS.getId() };

    const DriveApp = {
      getFolderById(id) { if (!folders[id]) throw new Error('No folder ' + id); return folders[id]; },
      getFileById(id) { if (!files[id]) throw new Error('No file ' + id); return files[id]; },
      getRootFolder() { return rootFolder; },
      _files: files,
      _folders: folders,
    };

    // ---------------- Utilities (incl. sync SHA-256 / HMAC) ----------------
    function utf8(str) {
      if (typeof TextEncoder !== 'undefined') return Array.from(new TextEncoder().encode(str));
      return Array.from(Buffer.from(str, 'utf8'));
    }
    function sha256(bytes) {
      const K = [0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
      let h = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
      const msg = bytes.slice();
      const bitLen = bytes.length * 8;
      msg.push(0x80);
      while (msg.length % 64 !== 56) msg.push(0);
      for (let i = 7; i >= 0; i--) msg.push(Math.floor(bitLen / Math.pow(2, i * 8)) & 0xff);
      const w = new Array(64);
      const rotr = (x, n) => (x >>> n) | (x << (32 - n));
      for (let off = 0; off < msg.length; off += 64) {
        for (let i = 0; i < 16; i++) w[i] = (msg[off + i * 4] << 24) | (msg[off + i * 4 + 1] << 16) | (msg[off + i * 4 + 2] << 8) | msg[off + i * 4 + 3];
        for (let i = 16; i < 64; i++) {
          const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
          const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
          w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0;
        }
        let [a, b, c, d, e, f, g, hh] = h;
        for (let i = 0; i < 64; i++) {
          const S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
          const ch = (e & f) ^ (~e & g);
          const S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
          const maj = (a & b) ^ (a & c) ^ (b & c);
          const t1 = (hh + S1 + ch + K[i] + w[i]) | 0;
          const t2 = (S0 + maj) | 0;
          hh = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
        }
        h = [(h[0] + a) | 0, (h[1] + b) | 0, (h[2] + c) | 0, (h[3] + d) | 0, (h[4] + e) | 0, (h[5] + f) | 0, (h[6] + g) | 0, (h[7] + hh) | 0];
      }
      const out = [];
      h.forEach((x) => { out.push((x >>> 24) & 255, (x >>> 16) & 255, (x >>> 8) & 255, x & 255); });
      return out;
    }
    function hmacSha256(msgBytes, keyBytes) {
      let key = keyBytes.length > 64 ? sha256(keyBytes) : keyBytes.slice();
      while (key.length < 64) key.push(0);
      const ipad = key.map((b) => b ^ 0x36), opad = key.map((b) => b ^ 0x5c);
      return sha256(opad.concat(sha256(ipad.concat(msgBytes))));
    }
    function toSigned(bytes) { return bytes.map((b) => (b > 127 ? b - 256 : b)); }
    function toUnsigned(bytes) { return bytes.map((b) => (b < 0 ? b + 256 : b)); }
    function b64encode(bytes) {
      const u = toUnsigned(bytes);
      if (typeof btoa !== 'undefined') { let s = ''; u.forEach((b) => { s += String.fromCharCode(b); }); return btoa(s); }
      return Buffer.from(u).toString('base64');
    }
    function b64decode(str) {
      if (typeof atob !== 'undefined') { const s = atob(str); const out = []; for (let i = 0; i < s.length; i++) out.push(s.charCodeAt(i)); return toSigned(out); }
      return toSigned(Array.from(Buffer.from(str, 'base64')));
    }
    function pad(n, w) { return String(n).padStart(w || 2, '0'); }
    const Utilities = {
      Charset: { UTF_8: 'UTF-8' },
      getUuid() {
        uuidSeq++;
        const hex = (uuidSeq.toString(16) + Math.random().toString(16).slice(2) + '0000000000000000').slice(0, 32);
        return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}`;
      },
      formatDate(date, tz, fmt) {
        const offset = tz === 'Asia/Tokyo' ? 9 * 3600000 : 0;
        const d = new Date(date.getTime() + offset);
        return fmt.replace('yyyy', d.getUTCFullYear()).replace('yy', String(d.getUTCFullYear()).slice(2))
          .replace('MM', pad(d.getUTCMonth() + 1)).replace('dd', pad(d.getUTCDate()))
          .replace('HH', pad(d.getUTCHours())).replace('mm', pad(d.getUTCMinutes())).replace('ss', pad(d.getUTCSeconds()));
      },
      base64Decode: b64decode,
      base64Encode(bytes) { return b64encode(bytes); },
      base64EncodeWebSafe(bytes) { return b64encode(bytes).replace(/\+/g, '-').replace(/\//g, '_'); },
      computeHmacSha256Signature(value, key) { return toSigned(hmacSha256(utf8(String(value)), utf8(String(key)))); },
      newBlob(data, type, name) { return { data, type, getName: () => name, getContentType: () => type, getBytes: () => data }; },
      sleep() {},
    };

    // ---------------- Cache / Properties / Lock ----------------
    const cacheStore = {};
    const CacheService = {
      getScriptCache() {
        return {
          get: (k) => (k in cacheStore ? cacheStore[k] : null),
          put: (k, v) => { if (String(v).length > 100000) throw new Error('too large'); cacheStore[k] = String(v); },
          remove: (k) => { delete cacheStore[k]; },
        };
      },
      _store: cacheStore,
    };
    const props = {};
    const PropertiesService = {
      getScriptProperties() {
        return {
          getProperty: (k) => (k in props ? props[k] : null),
          setProperty: (k, v) => { props[k] = String(v); },
          deleteProperty: (k) => { delete props[k]; },
          getProperties: () => Object.assign({}, props),
        };
      },
      _props: props,
    };
    let locked = false;
    const LockService = {
      getScriptLock() {
        return {
          tryLock: () => { log.lockCalls++; if (opts.lockFails) return false; locked = true; return true; },
          waitLock: () => { locked = true; },
          releaseLock: () => { locked = false; },
          hasLock: () => locked,
        };
      },
    };

    // ---------------- UI / Html / Content / Script ----------------
    const Button = { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO', CLOSE: 'CLOSE' };
    const ui = {
      Button,
      ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' },
      alert(title, msg) { log.alerts.push({ title, msg }); return opts.alertAnswer || Button.OK; },
      prompt(title, msg) {
        log.prompts.push({ title, msg });
        const a = promptAnswers.shift() || { button: Button.CANCEL, text: '' };
        return { getSelectedButton: () => a.button, getResponseText: () => a.text };
      },
      showModalDialog(html, title) { log.dialogs.push({ html, title }); },
      createMenu() { const m = { addItem: () => m, addSeparator: () => m, addToUi: () => m }; return m; },
    };
    function builder(kind) {
      const b = { kind, calls: [] };
      return new Proxy(b, {
        get(t, prop) {
          if (prop === 'build') return () => ({ kind, calls: t.calls });
          if (prop in t) return t[prop];
          return (...args) => { t.calls.push([prop, args]); return proxyB; };
        },
      });
    }
    let proxyB;
    const SpreadsheetApp = {
      getActiveSpreadsheet: () => activeSS,
      openById(id) { if (!spreadsheets[id]) throw new Error('No spreadsheet ' + id); return spreadsheets[id]; },
      openByUrl(url) {
        const m = /\/d\/([^/]+)/.exec(url || '');
        if (!m || !spreadsheets[m[1]]) throw new Error('No spreadsheet at ' + url);
        return spreadsheets[m[1]];
      },
      create(name) { return makeSpreadsheet('SS_' + (++fileSeq), name); },
      flush() {},
      getUi: () => ui,
      newDataValidation() { proxyB = builder('validation'); return proxyB; },
      newConditionalFormatRule() { proxyB = builder('cf'); return proxyB; },
      _spreadsheets: spreadsheets,
      _makeSpreadsheet: makeSpreadsheet,
    };
    function output(content) {
      const o = {
        content, title: '', meta: [], mime: 'text/plain',
        setTitle(t) { this.title = t; return this; },
        addMetaTag(n, c) { this.meta.push([n, c]); return this; },
        setXFrameOptionsMode() { return this; },
        setWidth() { return this; },
        setHeight() { return this; },
        setMimeType(m) { this.mime = m; return this; },
        getContent() { return this.content; },
      };
      return o;
    }
    const HtmlService = {
      XFrameOptionsMode: { DEFAULT: 'DEFAULT', ALLOWALL: 'ALLOWALL' },
      createHtmlOutputFromFile(name) { return output('[file:' + name + ']'); },
      createHtmlOutput(html) { return output(html); },
    };
    const ContentService = { MimeType: { JSON: 'application/json', TEXT: 'text/plain' }, createTextOutput: (s) => output(s) };
    const ScriptApp = { getService: () => ({ getUrl: () => opts.serviceUrl || 'https://script.google.com/macros/s/TESTDEPLOY/exec' }) };
    const Session = { getActiveUser: () => ({ getEmail: () => '' }), getScriptTimeZone: () => 'Asia/Tokyo' };

    return {
      globals: {
        SpreadsheetApp, DriveApp, Utilities, CacheService, PropertiesService, LockService,
        HtmlService, ContentService, ScriptApp, Session,
        console: opts.console || console,
      },
      activeSS, log, cacheStore, props, files, folders,
    };
  }

  const SERVER_EXPORTS = ['api', 'doGet', 'doPost', 'onOpen', 'onEdit', 'setup', 'runImport_', 'runLedgerCheck_',
    'inferImportStatus_', 'loadSettings_', 'getSettings_', 'handleRequest_', 'normalizeSerial_', 'importSummaryText_',
    'buildShareHtml_', 'importFromSheet', 'importFromSpreadsheetUrl', 'checkLedger', 'showShareDialog', 'makeToken_'];

  /** Node/test only: evaluate server sources with the mock globals */
  function loadServer(env, sources) {
    const names = Object.keys(env.globals);
    const body = sources.join('\n;\n') + '\n;return {' + SERVER_EXPORTS.join(',') + '};';
    // eslint-disable-next-line no-new-func
    const fn = new Function(...names, body);
    return fn(...names.map((n) => env.globals[n]));
  }

  const api = { createGasEnv, loadServer, SERVER_EXPORTS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.GasMock = api;
})(typeof window !== 'undefined' ? window : globalThis);


window.__makeServer = function (env) {
  var G = env.globals;
  return (function (SpreadsheetApp, DriveApp, Utilities, CacheService, PropertiesService, LockService, HtmlService, ContentService, ScriptApp, Session, console) {
/**
 * ============================================================
 *  impactTV 端末管理  ― サーバー側（Code.gs）  v1.0.0
 * ------------------------------------------------------------
 *  端末の黄色シール（バーコード）を読み取り、
 *  状態・設置場所・履歴をこのスプレッドシートに記録します。
 *
 *  ■ はじめに
 *   1. メニュー「impactTV管理」→「初期設定（シート作成・更新）」を実行
 *   2. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *        次のユーザーとして実行 ： 自分
 *        アクセスできるユーザー ： 全員
 *   3. 合言葉は「設定」シートで確認・変更できます
 * ============================================================
 */

const APP_VERSION = '1.0.0';
const TZ = 'Asia/Tokyo';
const SETTINGS_CACHE_KEY = 'itv:settings:v1';

const SHEET = {
  DEVICES: '台帳',
  LOG: '履歴',
  SUMMARY: '集計',
  IMPORT: '取込',
  STORES: '店舗マスタ',
  OPTIONS: '選択肢',
  SETTINGS: '設定',
};

/** 台帳の列（見出し名で読み書きするので、列の並べ替え・列の追加をしても動きます） */
const D = {
  serial: 'シリアル番号', location: '設置場所', spot: '設置個所', status: '状態', symptom: '症状',
  model: '機種', contract: '契約', memo: 'メモ', checkedAt: '最終確認日時', checkedBy: '最終確認者',
  updatedAt: '更新日時', updatedBy: '更新者', photo: '写真', createdAt: '登録日時', rev: '版', lastOp: '最終操作ID',
};
const DEVICE_FIELDS = Object.keys(D);
const DEVICE_HEADERS = DEVICE_FIELDS.map(function (k) { return D[k]; });
const DEVICE_DATE_FIELDS = ['checkedAt', 'updatedAt', 'createdAt'];

/** 履歴の列 */
const L = {
  time: '日時', opId: '操作ID', serial: 'シリアル番号', op: '操作', st0: '状態(前)', st1: '状態(後)',
  l0: '場所(前)', l1: '場所(後)', spot: '設置個所', symptom: '症状', memo: 'コメント', photo: '写真',
  user: '担当者', device: '端末', undone: '取消', snap: '復元用',
};
const LOG_FIELDS = Object.keys(L);
const LOG_HEADERS = LOG_FIELDS.map(function (k) { return L[k]; });

const STORE_HEADERS = ['店舗名', '区分', '表示順', '非表示', '備考'];
const IMPORT_HEADERS = ['端末番号(黄色シール)', '店舗名', '設置個所', '設置個所”その他”の場合どこに置いているか入力', '解約', 'メッセージ', '取込結果'];
const STATUS_KINDS = ['正常', '故障', '修理', '保管', '移動', '不明', '終了'];

/** 設定シートの項目 [キー, 項目名, 初期値, 説明] */
const SETTING_DEFS = [
  ['passcode', '合言葉', '', 'アプリを開くときに入力します。変更すると、全員にもう一度入力してもらう形になります。'],
  ['appName', 'アプリ名', 'impactTV 端末管理', '画面の上に表示される名前'],
  ['appUrl', 'アプリURL', '', 'デプロイ後の「ウェブアプリのURL（…/exec）」を貼ると、メニューから配布用QRコードを表示できます'],
  ['serialPattern', 'シリアル形式', '^[A-Z0-9]{6}[0-9]{7}$', '新規登録のとき、この形式と違えば確認メッセージを出します（空欄＝確認しない）'],
  ['undoMinutes', '取り消しできる時間（分）', 30, '保存後、この時間内なら「元に戻す」が使えます'],
  ['staleDays', '長期未確認（日）', 90, 'この日数以上読み取られていない端末を「長期未確認」として表示します'],
  ['homeLocation', '本部の保管場所', '本部在庫', '取込で「本部に返却」「本部在庫」と書かれた端末の設置場所'],
  ['formats', '読み取るバーコード', 'code_128,code_39,code_93,ean_13,ean_8,qr_code', '読み取る種類（黄色シールは code_128）'],
  ['photoFolderId', '写真フォルダID', '', '故障写真の保存先フォルダ（初期設定で自動作成）'],
];

/* ============================================================
 *  メニュー・ウェブアプリの入口
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('impactTV管理')
    .addItem('初期設定（シート作成・更新）', 'setup')
    .addSeparator()
    .addItem('「取込」シートの内容を台帳に取り込む', 'importFromSheet')
    .addItem('別のスプレッドシートから取り込む…', 'importFromSpreadsheetUrl')
    .addSeparator()
    .addItem('台帳のチェック（重複・入力ミス）', 'checkLedger')
    .addItem('配布用URL・QRコードを表示', 'showShareDialog')
    .addToUi();
}

/** マスタや設定が編集されたら、アプリ側のキャッシュを捨てる */
function onEdit(e) {
  try {
    const name = e && e.range && e.range.getSheet().getName();
    if ([SHEET.SETTINGS, SHEET.OPTIONS, SHEET.STORES].indexOf(name) >= 0) {
      CacheService.getScriptCache().remove(SETTINGS_CACHE_KEY);
    }
  } catch (err) { /* 何もしない */ }
}

function doGet(e) {
  let title = 'impactTV 端末管理';
  try { title = getSettings_().appName || title; } catch (err) { /* 初期設定前でも画面は出す */ }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** 外部ページ（GitHub Pages など）から呼ばれるとき */
function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_(fail_('BAD_REQUEST', 'リクエストの形式が正しくありません。'));
  }
  return json_(handleRequest_(req));
}

/** GAS内の画面（google.script.run）から呼ばれるとき */
function api(req) {
  return handleRequest_(req);
}

const ACTIONS = {
  login: { auth: false, fn: function (p, ctx) { return apiLogin_(p, ctx); } },
  bootstrap: { auth: true, fn: function (p, ctx) { return apiBootstrap_(p, ctx); } },
  getDevice: { auth: true, fn: function (p, ctx) { return apiGetDevice_(p, ctx); } },
  saveDevice: { auth: true, idem: true, fn: function (p, ctx) { return apiSaveDevice_(p, ctx); } },
  undo: { auth: true, idem: true, fn: function (p, ctx) { return apiUndo_(p, ctx); } },
  inventoryCommit: { auth: true, idem: true, fn: function (p, ctx) { return apiInventoryCommit_(p, ctx); } },
};

function handleRequest_(req) {
  try {
    req = req || {};
    const def = ACTIONS[req.action];
    if (!def) return fail_('BAD_REQUEST', '不明な操作です。');
    const ctx = {
      settings: getSettings_(),
      user: cleanText_(req.user, 40),
      device: sanitizeClient_(req.device),
    };
    if (def.auth && !verifyToken_(req.token, ctx.settings)) {
      return fail_('AUTH', '合言葉の確認が必要です。もう一度入力してください。');
    }
    const payload = req.payload || {};
    // 同じ操作の二重送信（通信の再送など）は、1回目の結果をそのまま返す
    let cacheKey = null;
    if (def.idem && payload.clientOpId) {
      cacheKey = 'itv:op:' + String(payload.clientOpId).slice(0, 64);
      const hit = CacheService.getScriptCache().get(cacheKey);
      if (hit) {
        const cached = JSON.parse(hit);
        cached.replayed = true;
        return cached;
      }
    }
    const data = def.fn(payload, ctx);
    const res = { ok: true, data: data, serverTime: Date.now(), version: APP_VERSION };
    if (cacheKey) {
      try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(res), 21600); } catch (err) { /* 大きすぎる場合は保存しない */ }
    }
    return res;
  } catch (err) {
    if (err && err.appCode) return fail_(err.appCode, err.message, err.extra);
    console.error(err && err.stack ? err.stack : err);
    return fail_('SERVER', 'エラーが発生しました：' + (err && err.message ? err.message : err));
  }
}

/* ============================================================
 *  API
 * ============================================================ */

function apiLogin_(p, ctx) {
  const S = ctx.settings;
  if (!S.passcode) throw appError_('SETUP', '合言葉が設定されていません。本部に連絡してください。');
  const cache = CacheService.getScriptCache();
  const devKey = 'itv:fail:' + (ctx.device.id || 'unknown');
  const allKey = 'itv:fail:all';
  const devFails = Number(cache.get(devKey) || 0);
  const allFails = Number(cache.get(allKey) || 0);
  if (devFails >= 10 || allFails >= 100) {
    throw appError_('LOCKED', '合言葉の入力ミスが続いたため、しばらく入力できません。15分ほど待ってから試してください。');
  }
  if (normalizePass_(p.passcode) !== normalizePass_(S.passcode)) {
    cache.put(devKey, String(devFails + 1), 900);
    cache.put(allKey, String(allFails + 1), 900);
    throw appError_('BAD_PASS', '合言葉が違います。');
  }
  cache.remove(devKey);
  return { token: makeToken_(S.passcode) };
}

function apiBootstrap_(p, ctx) {
  const S = ctx.settings;
  const t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
  const devices = [];
  const seen = {};
  const dupSerials = [];
  t.rows.forEach(function (row, i) {
    const d = rowToDevice_(t, row, i + 2);
    if (!d.serial) return;
    if (seen[d.serial]) { if (dupSerials.indexOf(d.serial) < 0) dupSerials.push(d.serial); return; }
    seen[d.serial] = true;
    devices.push(compactDevice_(d));
  });
  return {
    settings: {
      appName: S.appName, serialPattern: S.serialPattern, undoMinutes: S.undoMinutes,
      staleDays: S.staleDays, homeLocation: S.homeLocation, formats: S.formats,
    },
    statuses: S.statuses, spots: S.spots, symptoms: S.symptoms, models: S.models, contracts: S.contracts,
    stores: S.stores,
    devices: devices,
    duplicates: dupSerials,
  };
}

function apiGetDevice_(p, ctx) {
  const serial = requireSerial_(p.serial);
  const t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
  const found = findDevice_(t, serial);
  return {
    serial: serial,
    device: found ? publicDevice_(found) : null,
    history: readHistory_(serial, 15),
  };
}

/**
 * 端末の保存
 *  mode: 'new'（新規登録） / 'update'（状態・場所などの変更） / 'check'（確認のみ）
 */
function apiSaveDevice_(p, ctx) {
  requireUser_(ctx);
  const S = ctx.settings;
  const serial = requireSerial_(p.serial);
  const mode = ['new', 'update', 'check'].indexOf(p.mode) >= 0 ? p.mode : 'update';

  // 1) ロック前の事前チェック（写真を保存してから衝突に気づくのを減らす）
  let t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
  const pre = judgeSave_(mode, findDevice_(t, serial), p, serial);
  if (pre) return pre;

  // 2) 写真の保存（時間がかかるのでロックの外で）
  const photos = mode === 'check' ? [] : savePhotos_(serial, p.photos, ctx);

  // 3) ロックして確定
  return withLock_(function () {
    t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
    const found = findDevice_(t, serial);
    const again = judgeSave_(mode, found, p, serial);
    if (again) { trashFiles_(photos); return again; }

    const now = new Date();
    const opId = newOpId_();
    const user = ctx.user;
    const client = clientLabel_(ctx);
    const comment = cleanText_(p.comment, 500);
    const photoText = photos.map(function (f) { return f.url; }).join('\n');

    if (mode === 'new') {
      const status = p.status ? requireStatus_(S, p.status) : defaultStatus_(S);
      const location = requireLocation_(S, p.location);
      const spot = p.spot ? requireSpot_(S, p.spot) : '';
      const symptom = statusKind_(S, status) === '故障' ? buildSymptom_(p, comment) : '';
      const dev = {
        serial: serial, location: location, spot: spot, status: status, symptom: symptom,
        model: cleanText_(p.model, 40) || inferModel_(S, serial), contract: '',
        memo: cleanText_(p.deviceMemo, 500), checkedAt: now, checkedBy: user, updatedAt: now, updatedBy: user,
        photo: photoText, createdAt: now, rev: 1, lastOp: opId,
      };
      appendDevices_(t, [dev]);
      appendLogs_([{
        time: now, opId: opId, serial: serial, op: '新規登録', st0: '', st1: status, l0: '', l1: location,
        spot: spot, symptom: symptom, memo: comment, photo: photoText, user: user, device: client, undone: '', snap: 'NEW',
      }]);
      return { result: 'saved', op: '新規登録', opId: opId, undoable: true, device: publicDevice_(dev) };
    }

    const before = snapshot_(found);

    if (mode === 'check') {
      found.checkedAt = now;
      found.checkedBy = user;
      writeDevice_(t, found);
      appendLogs_([{
        time: now, opId: opId, serial: serial, op: '確認', st0: found.status, st1: found.status,
        l0: found.location, l1: found.location, spot: found.spot, symptom: '', memo: comment, photo: '',
        user: user, device: client, undone: '', snap: '',
      }]);
      return { result: 'saved', op: '確認', opId: opId, undoable: false, device: publicDevice_(found) };
    }

    // update
    const newStatus = (p.status && p.status !== found.status) ? requireStatus_(S, p.status) : found.status;
    const newLocation = (p.location && normKey_(p.location) !== normKey_(found.location)) ? requireLocation_(S, p.location) : found.location;
    let newSpot = found.spot;
    if (p.spot != null && String(p.spot) !== found.spot) newSpot = p.spot ? requireSpot_(S, p.spot) : '';
    const kind = statusKind_(S, newStatus);
    let newSymptom = '';
    if (kind === '故障') newSymptom = buildSymptom_(p, comment) || found.symptom;
    const newMemo = p.deviceMemo != null ? cleanText_(p.deviceMemo, 500) : found.memo;

    const statusChanged = newStatus !== found.status;
    const locChanged = newLocation !== found.location;
    let op;
    if (statusChanged && locChanged) op = '移動・状態変更';
    else if (locChanged) op = '移動';
    else if (statusChanged) op = kind === '故障' ? '故障報告' : '状態変更';
    else if (kind === '故障' && p.kind === 'report') op = '故障報告（追記）';
    else if (newSpot !== found.spot || newMemo !== found.memo || newSymptom !== found.symptom) op = '内容の変更';
    else if (photos.length || comment) op = 'コメント・写真';
    else op = '確認';

    found.status = newStatus;
    found.location = newLocation;
    found.spot = newSpot;
    found.symptom = newSymptom;
    found.memo = newMemo;
    if (photoText) found.photo = photoText;
    found.checkedAt = now;
    found.checkedBy = user;
    const undoable = op !== '確認';
    if (undoable) {
      found.updatedAt = now;
      found.updatedBy = user;
      found.rev = (found.rev || 0) + 1;
      found.lastOp = opId;
    }
    writeDevice_(t, found);
    appendLogs_([{
      time: now, opId: opId, serial: serial, op: op, st0: before.status, st1: newStatus,
      l0: before.location, l1: newLocation, spot: newSpot, symptom: newSymptom, memo: comment,
      photo: photoText, user: user, device: client, undone: '', snap: undoable ? JSON.stringify(before) : '',
    }]);
    return { result: 'saved', op: op, opId: opId, undoable: undoable, device: publicDevice_(found) };
  });
}

/** 保存前の判定（登録済み・未登録・他の人が先に更新） */
function judgeSave_(mode, found, p, serial) {
  if (mode === 'new' && found) return { result: 'exists', device: publicDevice_(found) };
  if (mode !== 'new' && !found) return { result: 'notfound', serial: serial };
  if (mode === 'update' && Number(p.baseRev) !== Number(found.rev || 0)) {
    return { result: 'conflict', device: publicDevice_(found), last: readHistory_(found.serial, 1)[0] || null };
  }
  return null;
}

/** 直前の操作を取り消す */
function apiUndo_(p, ctx) {
  requireUser_(ctx);
  const S = ctx.settings;
  const opId = cleanText_(p.opId, 40);
  if (!opId) throw appError_('BAD_REQUEST', '取り消す操作が指定されていません。');
  return withLock_(function () {
    const log = openTable_(SHEET.LOG, LOG_HEADERS);
    const rows = findRowsByValue_(log, L.opId, opId).filter(function (r) {
      return !r.values[log.idx[L.undone]] && r.values[log.idx[L.snap]];
    });
    if (!rows.length) throw appError_('UNDO_NOT_FOUND', '取り消せる操作が見つかりません（すでに取り消し済みか、確認だけの操作です）。');
    const now = new Date();
    const oldest = Math.min.apply(null, rows.map(function (r) { return toMs_(r.values[log.idx[L.time]]) || 0; }));
    if (now.getTime() - oldest > S.undoMinutes * 60000) {
      throw appError_('UNDO_EXPIRED', '保存から' + S.undoMinutes + '分を過ぎたため、取り消せません。もう一度内容を変更して保存してください。');
    }
    const t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
    const undoId = newOpId_();
    const client = clientLabel_(ctx);
    const items = [];
    const logs = [];
    const deleteRows = [];
    rows.forEach(function (r) {
      const serial = normalizeSerial_(r.values[log.idx[L.serial]]);
      const snap = String(r.values[log.idx[L.snap]]);
      let dev;
      try { dev = findDevice_(t, serial); } catch (err) { items.push({ serial: serial, ok: false, reason: err.message }); return; }
      if (!dev) { items.push({ serial: serial, ok: false, reason: '台帳にありません' }); return; }
      if (dev.lastOp !== opId) { items.push({ serial: serial, ok: false, reason: 'その後に別の操作があったため' }); return; }
      if (snap === 'NEW') {
        deleteRows.push(dev.row);
        logs.push({
          time: now, opId: undoId, serial: serial, op: '取り消し（登録を削除）', st0: dev.status, st1: '',
          l0: dev.location, l1: '', spot: '', symptom: '', memo: '元の操作ID ' + opId, photo: '',
          user: ctx.user, device: client, undone: '', snap: '',
        });
      } else {
        const before = parseSnapshot_(snap);
        const restored = Object.assign({}, dev, before, {
          row: dev.row, rev: (dev.rev || 0) + 1, lastOp: undoId, updatedAt: now, updatedBy: ctx.user,
        });
        writeDevice_(t, restored);
        logs.push({
          time: now, opId: undoId, serial: serial, op: '取り消し', st0: dev.status, st1: restored.status,
          l0: dev.location, l1: restored.location, spot: restored.spot, symptom: restored.symptom,
          memo: '元の操作ID ' + opId, photo: '', user: ctx.user, device: client, undone: '', snap: '',
        });
      }
      items.push({ serial: serial, ok: true });
    });
    const okCount = items.filter(function (x) { return x.ok; }).length;
    if (okCount) {
      const mark = '取消済 ' + fmtDate_(now) + ' ' + ctx.user;
      rows.forEach(function (r) {
        log.sheet.getRange(r.row, log.idx[L.undone] + 1).setValue(mark);
      });
      deleteRows.sort(function (a, b) { return b - a; }).forEach(function (rowNo) { t.sheet.deleteRow(rowNo); });
      appendLogs_(logs);
    }
    return { result: okCount ? 'undone' : 'failed', undoId: undoId, items: items, okCount: okCount, ngCount: items.length - okCount };
  });
}

/**
 * 棚卸の確定
 *  scanned  : 読み取ったシリアル
 *  move     : 別の場所で登録されていた端末のうち「この場所へ移動」とするもの
 *  register : 未登録の端末のうち「この場所で登録」するもの
 *  missing  : この場所の台帳にあるが見つからなかった端末のうち「所在不明」にするもの
 *  revive   : 所在不明・移動中だった端末のうち「稼働中」に戻すもの
 */
function apiInventoryCommit_(p, ctx) {
  requireUser_(ctx);
  const S = ctx.settings;
  const location = requireLocation_(S, p.location);
  const list = function (a) { return uniq_((a || []).map(normalizeSerial_).filter(Boolean)); };
  const scanned = list(p.scanned).slice(0, 1000);
  const move = toSet_(list(p.move));
  const register = toSet_(list(p.register));
  const missing = list(p.missing);
  const revive = toSet_(list(p.revive));
  const scannedSet = toSet_(scanned);
  const normal = defaultStatus_(S);
  const lost = kindStatus_(S, '不明');

  return withLock_(function () {
    const t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
    const index = indexDevices_(t);
    const now = new Date();
    const opId = newOpId_();
    const user = ctx.user;
    const client = clientLabel_(ctx);
    const counts = { checked: 0, moved: 0, registered: 0, missing: 0, revived: 0, skipped: 0 };
    const warnings = [];
    const logs = [];
    const newDevices = [];
    const logOf = function (dev, before, op, snap) {
      return {
        time: now, opId: opId, serial: dev.serial, op: op, st0: before ? before.status : '', st1: dev.status,
        l0: before ? before.location : '', l1: dev.location, spot: dev.spot, symptom: dev.symptom,
        memo: '棚卸：' + location, photo: '', user: user, device: client, undone: '', snap: snap,
      };
    };

    scanned.forEach(function (serial) {
      const hits = index[serial];
      if (hits && hits.length > 1) {
        warnings.push(serial + '：台帳に' + hits.length + '行あるため更新しませんでした');
        counts.skipped++;
        return;
      }
      if (!hits) {
        if (!register[serial]) { counts.skipped++; return; }
        if (!/^[A-Z0-9\-_.\/+]{1,40}$/.test(serial)) { warnings.push(serial + '：形式が正しくないため登録しませんでした'); counts.skipped++; return; }
        const dev = {
          serial: serial, location: location, spot: '', status: normal, symptom: '', model: inferModel_(S, serial),
          contract: '', memo: '', checkedAt: now, checkedBy: user, updatedAt: now, updatedBy: user, photo: '',
          createdAt: now, rev: 1, lastOp: opId,
        };
        newDevices.push(dev);
        logs.push(logOf(dev, null, '新規登録（棚卸）', 'NEW'));
        counts.registered++;
        return;
      }
      const dev = rowToDevice_(t, t.rows[hits[0]], hits[0] + 2);
      const before = snapshot_(dev);
      let changed = false;
      let op = '棚卸確認';
      if (normKey_(dev.location) !== normKey_(location) && move[serial]) {
        dev.location = location;
        changed = true;
        op = '移動（棚卸）';
        counts.moved++;
      }
      const kind = statusKind_(S, dev.status);
      if ((kind === '不明' || kind === '移動') && revive[serial] && dev.status !== normal) {
        dev.status = normal;
        dev.symptom = '';
        op = changed ? '移動・状態変更（棚卸）' : '状態変更（棚卸）';
        changed = true;
        counts.revived++;
      }
      dev.checkedAt = now;
      dev.checkedBy = user;
      if (changed) {
        dev.rev = (dev.rev || 0) + 1;
        dev.lastOp = opId;
        dev.updatedAt = now;
        dev.updatedBy = user;
      }
      writeDevice_(t, dev);
      logs.push(logOf(dev, before, op, changed ? JSON.stringify(before) : ''));
      counts.checked++;
    });

    if (lost) {
      missing.forEach(function (serial) {
        if (scannedSet[serial]) return;
        const hits = index[serial];
        if (!hits || hits.length !== 1) return;
        const dev = rowToDevice_(t, t.rows[hits[0]], hits[0] + 2);
        if (normKey_(dev.location) !== normKey_(location) || dev.status === lost) return;
        const before = snapshot_(dev);
        dev.status = lost;
        dev.symptom = '';
        dev.rev = (dev.rev || 0) + 1;
        dev.lastOp = opId;
        dev.updatedAt = now;
        dev.updatedBy = user;
        writeDevice_(t, dev);
        logs.push(logOf(dev, before, '所在不明（棚卸）', JSON.stringify(before)));
        counts.missing++;
      });
    }

    appendDevices_(t, newDevices);
    appendLogs_(logs);
    const changes = counts.moved + counts.registered + counts.missing + counts.revived;
    return { result: 'saved', opId: opId, counts: counts, warnings: warnings, undoable: changes > 0 };
  });
}

/* ============================================================
 *  台帳の読み書き
 * ============================================================ */

function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw appError_('SETUP', '初期設定がまだです。スプレッドシートのメニュー「impactTV管理」→「初期設定」を実行してください。');
  return ss;
}

function getSheet_(name) {
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw appError_('SETUP', '「' + name + '」シートがありません。メニュー「impactTV管理」→「初期設定」を実行してください。');
  return sh;
}

/** シート全体を見出しつきで読む（台帳など、行数が多すぎないシート用） */
function readTable_(name, required) {
  const sh = getSheet_(name);
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const range = sh.getRange(1, 1, Math.max(lastRow, 1), lastCol);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const t = makeHeader_(sh, values[0], required);
  t.rows = values.slice(1);
  t.formulas = formulas.slice(1);
  t.lastRow = lastRow;
  return t;
}

/** 見出しだけ読む（履歴のように行数が増えるシート用） */
function openTable_(name, required) {
  const sh = getSheet_(name);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const t = makeHeader_(sh, sh.getRange(1, 1, 1, lastCol).getValues()[0], required);
  t.lastRow = sh.getLastRow();
  return t;
}

function makeHeader_(sh, headerRow, required) {
  const headers = headerRow.map(function (h) { return String(h).trim(); });
  const idx = {};
  headers.forEach(function (h, i) { if (h && !(h in idx)) idx[h] = i; });
  (required || []).forEach(function (h) {
    if (!(h in idx)) throw appError_('SETUP', '「' + sh.getName() + '」シートに見出し「' + h + '」がありません。メニュー「impactTV管理」→「初期設定」を実行してください。');
  });
  return { sheet: sh, headers: headers, idx: idx, width: headers.length };
}

function rowToDevice_(t, row, rowNumber) {
  const d = { row: rowNumber };
  DEVICE_FIELDS.forEach(function (f) {
    const v = row[t.idx[D[f]]];
    if (DEVICE_DATE_FIELDS.indexOf(f) >= 0) d[f] = v instanceof Date ? v : (v ? new Date(v) : '');
    else if (f === 'rev') d[f] = Number(v) || 0;
    else d[f] = v == null ? '' : String(v).trim();
  });
  d.serial = normalizeSerial_(d.serial);
  return d;
}

function deviceToRow_(t, dev, baseRow, baseFormulas) {
  const row = baseRow ? baseRow.slice() : [];
  while (row.length < t.width) row.push('');
  DEVICE_FIELDS.forEach(function (f) {
    let v = dev[f];
    if (v == null) v = '';
    if (DEVICE_DATE_FIELDS.indexOf(f) >= 0 && v !== '' && !(v instanceof Date)) v = new Date(v);
    row[t.idx[D[f]]] = v;
  });
  const out = row.map(function (v, i) { return baseFormulas && baseFormulas[i] ? baseFormulas[i] : sheetText_(v); });
  return { raw: row, out: out };
}

function findDevice_(t, serial) {
  const c = t.idx[D.serial];
  const hits = [];
  for (let i = 0; i < t.rows.length; i++) {
    if (normalizeSerial_(t.rows[i][c]) === serial) hits.push(i);
  }
  if (hits.length > 1) {
    throw appError_('DUPLICATE_ROWS', '台帳に「' + serial + '」が' + hits.length + '行あります（' +
      hits.map(function (i) { return (i + 2) + '行目'; }).join('、') + '）。本部で重複を整理してください。');
  }
  return hits.length ? rowToDevice_(t, t.rows[hits[0]], hits[0] + 2) : null;
}

function indexDevices_(t) {
  const c = t.idx[D.serial];
  const index = {};
  t.rows.forEach(function (r, i) {
    const s = normalizeSerial_(r[c]);
    if (!s) return;
    (index[s] = index[s] || []).push(i);
  });
  return index;
}

function writeDevice_(t, dev) {
  const i = dev.row - 2;
  const r = deviceToRow_(t, dev, t.rows[i], t.formulas && t.formulas[i]);
  t.sheet.getRange(dev.row, 1, 1, t.width).setValues([r.out]);
  t.rows[i] = r.raw;
}

function appendDevices_(t, devs) {
  if (!devs.length) return;
  const rows = devs.map(function (d) { return deviceToRow_(t, d, null, null); });
  const start = t.sheet.getLastRow() + 1;
  t.sheet.getRange(start, 1, rows.length, t.width).setValues(rows.map(function (r) { return r.out; }));
  rows.forEach(function (r, k) {
    devs[k].row = start + k;
    t.rows[start + k - 2] = r.raw;
  });
}

function appendLogs_(entries) {
  if (!entries.length) return;
  const t = openTable_(SHEET.LOG, LOG_HEADERS);
  const rows = entries.map(function (e) {
    const row = [];
    while (row.length < t.width) row.push('');
    LOG_FIELDS.forEach(function (f) { row[t.idx[L[f]]] = sheetText_(e[f] == null ? '' : e[f]); });
    return row;
  });
  t.sheet.getRange(t.sheet.getLastRow() + 1, 1, rows.length, t.width).setValues(rows);
}

/** 指定列が value と完全一致する行を探す（TextFinder で高速に） */
function findRowsByValue_(t, header, value) {
  const col = t.idx[header] + 1;
  if (t.lastRow < 2) return [];
  const cells = t.sheet.getRange(2, col, t.lastRow - 1, 1)
    .createTextFinder(String(value)).matchEntireCell(true).matchCase(true).findAll();
  return cells.map(function (c) {
    const r = c.getRow();
    return { row: r, values: t.sheet.getRange(r, 1, 1, t.width).getValues()[0] };
  });
}

function readHistory_(serial, limit) {
  const t = openTable_(SHEET.LOG, LOG_HEADERS);
  if (t.lastRow < 2) return [];
  const col = t.idx[L.serial] + 1;
  const cells = t.sheet.getRange(2, col, t.lastRow - 1, 1)
    .createTextFinder(serial).matchEntireCell(true).findAll();
  const rowsNo = cells.map(function (c) { return c.getRow(); }).sort(function (a, b) { return b - a; }).slice(0, limit);
  return rowsNo.map(function (r) {
    const v = t.sheet.getRange(r, 1, 1, t.width).getValues()[0];
    const g = function (k) { const x = v[t.idx[L[k]]]; return x == null ? '' : x; };
    return {
      t: toMs_(g('time')), id: String(g('opId')), op: String(g('op')),
      st0: String(g('st0')), st1: String(g('st1')), l0: String(g('l0')), l1: String(g('l1')),
      sy: String(g('symptom')), c: String(g('memo')), u: String(g('user')),
      ph: String(g('photo')).split('\n').filter(Boolean), undone: String(g('undone')),
    };
  });
}

function snapshot_(dev) {
  const s = {};
  DEVICE_FIELDS.forEach(function (f) {
    const v = dev[f];
    s[f] = v instanceof Date ? { $d: v.getTime() } : v;
  });
  return s;
}

function parseSnapshot_(text) {
  const s = JSON.parse(text);
  Object.keys(s).forEach(function (k) {
    if (s[k] && typeof s[k] === 'object' && '$d' in s[k]) s[k] = new Date(s[k].$d);
  });
  return s;
}

function publicDevice_(d) {
  return {
    serial: d.serial, location: d.location, spot: d.spot, status: d.status, symptom: d.symptom,
    model: d.model, contract: d.contract, memo: d.memo,
    checkedAt: toMs_(d.checkedAt), checkedBy: d.checkedBy, updatedAt: toMs_(d.updatedAt), updatedBy: d.updatedBy,
    createdAt: toMs_(d.createdAt), photos: String(d.photo || '').split('\n').filter(Boolean),
    rev: d.rev || 0, lastOp: d.lastOp || '',
  };
}

function compactDevice_(d) {
  return {
    s: d.serial, l: d.location, p: d.spot, st: d.status, sy: d.symptom, m: d.model, c: d.contract,
    n: d.memo, ck: toMs_(d.checkedAt), cb: d.checkedBy, up: toMs_(d.updatedAt), ub: d.updatedBy,
    r: d.rev || 0, ph: d.photo ? 1 : 0,
  };
}

/* ============================================================
 *  設定・マスタ
 * ============================================================ */

function getSettings_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(SETTINGS_CACHE_KEY);
  if (hit) return JSON.parse(hit);
  const s = loadSettings_();
  try { cache.put(SETTINGS_CACHE_KEY, JSON.stringify(s), 120); } catch (err) { /* 何もしない */ }
  return s;
}

function loadSettings_() {
  const s = {};
  SETTING_DEFS.forEach(function (d) { s[d[0]] = d[2]; });
  const values = getSheet_(SHEET.SETTINGS).getDataRange().getValues();
  const byLabel = {};
  values.forEach(function (r) { byLabel[String(r[0]).trim()] = r[1]; });
  SETTING_DEFS.forEach(function (d) {
    if (d[1] in byLabel && byLabel[d[1]] !== '' && byLabel[d[1]] != null) s[d[0]] = byLabel[d[1]];
  });
  s.passcode = String(s.passcode || '').trim();
  s.appName = String(s.appName || 'impactTV 端末管理');
  s.serialPattern = String(s.serialPattern || '').trim();
  s.undoMinutes = Math.max(1, Number(s.undoMinutes) || 30);
  s.staleDays = Math.max(1, Number(s.staleDays) || 90);
  s.homeLocation = String(s.homeLocation || '本部在庫').trim();
  s.formats = String(s.formats || '').split(/[,\s、]+/).map(function (x) { return x.trim().toLowerCase(); }).filter(Boolean);
  if (!s.formats.length) s.formats = ['code_128'];
  s.appUrl = String(s.appUrl || '').trim();
  s.photoFolderId = String(s.photoFolderId || '').trim();

  const opt = readOptions_(getSheet_(SHEET.OPTIONS));
  s.statuses = opt.statuses;
  s.spots = opt.spots;
  s.symptoms = opt.symptoms;
  s.models = opt.models;
  s.contracts = opt.contracts;
  s.stores = readStores_(getSheet_(SHEET.STORES));
  if (!s.statuses.length) throw appError_('SETUP', '「選択肢」シートに状態がありません。');
  return s;
}

function readOptions_(sh) {
  const values = sh.getDataRange().getValues();
  const head = (values[0] || []).map(function (v) { return String(v).trim(); });
  const col = function (name) { return head.indexOf(name); };
  const column = function (name) {
    const c = col(name);
    const out = [];
    if (c < 0) return out;
    for (let r = 1; r < values.length; r++) {
      const v = String(values[r][c]).trim();
      if (v && out.indexOf(v) < 0) out.push(v);
    }
    return out;
  };
  const statuses = [];
  const cs = col('状態'), ck = col('種別'), cc = col('色'), cd = col('説明');
  for (let r = 1; r < values.length; r++) {
    const name = cs >= 0 ? String(values[r][cs]).trim() : '';
    if (!name || statuses.some(function (x) { return x.name === name; })) continue;
    const kind = ck >= 0 ? String(values[r][ck]).trim() : '';
    const color = cc >= 0 ? String(values[r][cc]).trim() : '';
    statuses.push({
      name: name,
      kind: STATUS_KINDS.indexOf(kind) >= 0 ? kind : '',
      color: /^#[0-9a-f]{6}$/i.test(color) ? color : '#475569',
      desc: cd >= 0 ? String(values[r][cd]).trim() : '',
    });
  }
  const models = [];
  const mp = col('機種の先頭文字'), mn = col('機種名');
  for (let r = 1; r < values.length; r++) {
    const pre = mp >= 0 ? normalizeSerial_(values[r][mp]) : '';
    const nm = mn >= 0 ? String(values[r][mn]).trim() : '';
    if (pre && nm) models.push({ prefix: pre, name: nm });
  }
  return {
    statuses: statuses,
    spots: column('設置個所'),
    symptoms: column('症状'),
    models: models,
    contracts: column('契約'),
  };
}

function readStores_(sh) {
  const values = sh.getDataRange().getValues();
  const head = (values[0] || []).map(function (v) { return String(v).trim(); });
  const cName = head.indexOf('店舗名'), cGroup = head.indexOf('区分'), cOrder = head.indexOf('表示順'), cHide = head.indexOf('非表示');
  const list = [];
  const seen = {};
  for (let r = 1; r < values.length; r++) {
    const name = cName >= 0 ? String(values[r][cName]).trim() : '';
    if (!name || seen[normKey_(name)]) continue;
    const hide = cHide >= 0 ? values[r][cHide] : false;
    if (hide === true || String(hide).toUpperCase() === 'TRUE') continue;
    seen[normKey_(name)] = true;
    const rawOrder = cOrder >= 0 ? values[r][cOrder] : '';
    const order = rawOrder === '' || rawOrder == null ? 9999 : Number(rawOrder);
    list.push({ name: name, group: cGroup >= 0 ? String(values[r][cGroup]).trim() : '', order: isNaN(order) ? 9999 : order, i: r });
  }
  list.sort(function (a, b) { return a.order - b.order || a.i - b.i; });
  return list.map(function (x) { return { name: x.name, group: x.group }; });
}

function statusKind_(S, name) {
  for (let i = 0; i < S.statuses.length; i++) if (S.statuses[i].name === name) return S.statuses[i].kind;
  return '';
}

function kindStatus_(S, kind) {
  for (let i = 0; i < S.statuses.length; i++) if (S.statuses[i].kind === kind) return S.statuses[i].name;
  return null;
}

function defaultStatus_(S) {
  return kindStatus_(S, '正常') || S.statuses[0].name;
}

function inferModel_(S, serial) {
  let best = null;
  (S.models || []).forEach(function (m) {
    if (serial.indexOf(m.prefix) === 0 && (!best || m.prefix.length > best.prefix.length)) best = m;
  });
  return best ? best.name : '';
}

function requireStatus_(S, name) {
  name = cleanText_(name, 40);
  if (!S.statuses.some(function (x) { return x.name === name; })) {
    throw appError_('BAD_STATUS', '状態「' + name + '」は「選択肢」シートにありません。画面を更新してください。');
  }
  return name;
}

function requireLocation_(S, name) {
  const key = normKey_(name);
  if (!key) throw appError_('BAD_LOCATION', '設置場所を選んでください。');
  for (let i = 0; i < S.stores.length; i++) if (normKey_(S.stores[i].name) === key) return S.stores[i].name;
  throw appError_('BAD_LOCATION', '設置場所「' + cleanText_(name, 40) + '」は店舗マスタにありません。画面を更新してください。');
}

function requireSpot_(S, name) {
  name = cleanText_(name, 40);
  if (S.spots.length && S.spots.indexOf(name) < 0) {
    throw appError_('BAD_SPOT', '設置個所「' + name + '」は「選択肢」シートにありません。');
  }
  return name;
}

function requireSerial_(v) {
  const s = normalizeSerial_(v);
  if (!s) throw appError_('BAD_SERIAL', 'シリアル番号が空です。');
  if (s.length > 40 || !/^[A-Z0-9\-_.\/+]+$/.test(s)) {
    throw appError_('BAD_SERIAL', 'シリアル番号の形式が正しくありません（' + s.slice(0, 40) + '）。');
  }
  return s;
}

function requireUser_(ctx) {
  if (!ctx.user) throw appError_('NO_NAME', 'お名前を入力してください。');
}

function buildSymptom_(p, comment) {
  const list = (Array.isArray(p.symptoms) ? p.symptoms : []).map(function (x) { return cleanText_(x, 40); }).filter(Boolean).slice(0, 8);
  let text = list.join('、');
  if (comment) text = text ? text + '／' + comment : comment;
  return text.slice(0, 300);
}

/* ============================================================
 *  写真
 * ============================================================ */

function savePhotos_(serial, photos, ctx) {
  if (!Array.isArray(photos) || !photos.length) return [];
  const folder = getPhotoFolder_(ctx.settings);
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
  return photos.slice(0, 3).map(function (dataUrl, i) {
    const m = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
    if (!m) throw appError_('BAD_PHOTO', '写真の形式が正しくありません。');
    const bytes = Utilities.base64Decode(m[2]);
    if (bytes.length > 8 * 1024 * 1024) throw appError_('BAD_PHOTO', '写真が大きすぎます。');
    const ext = m[1] === 'image/png' ? 'png' : (m[1] === 'image/webp' ? 'webp' : 'jpg');
    const file = folder.createFile(Utilities.newBlob(bytes, m[1], serial + '_' + stamp + '_' + (i + 1) + '.' + ext));
    try { file.setDescription(serial + ' / ' + ctx.user + ' / ' + ctx.device.label); } catch (err) { /* 何もしない */ }
    return { id: file.getId(), url: file.getUrl() };
  });
}

function getPhotoFolder_(S) {
  if (S.photoFolderId) {
    try { return DriveApp.getFolderById(S.photoFolderId); } catch (err) { /* 下のエラーへ */ }
  }
  throw appError_('SETUP', '写真の保存先フォルダが見つかりません。メニュー「impactTV管理」→「初期設定」を実行してください。');
}

function trashFiles_(files) {
  (files || []).forEach(function (f) {
    try { DriveApp.getFileById(f.id).setTrashed(true); } catch (err) { /* 何もしない */ }
  });
}

/* ============================================================
 *  認証（合言葉 → トークン）
 * ============================================================ */

function normalizePass_(v) {
  return nfkc_(String(v == null ? '' : v)).trim().toLowerCase();
}

function makeToken_(passcode) {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('TOKEN_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('TOKEN_SECRET', secret);
  }
  const sig = Utilities.computeHmacSha256Signature(normalizePass_(passcode), secret, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '');
}

function verifyToken_(token, S) {
  if (!token || !S.passcode) return false;
  return String(token) === makeToken_(S.passcode);
}

/* ============================================================
 *  共通
 * ============================================================ */

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw appError_('BUSY', 'ただいま混み合っています。少し待ってからもう一度保存してください。');
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function appError_(code, message, extra) {
  const e = new Error(message);
  e.appCode = code;
  e.extra = extra;
  return e;
}

function fail_(code, message, extra) {
  return { ok: false, code: code, message: message, extra: extra || null, version: APP_VERSION };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clientLabel_(ctx) {
  return ctx.device.label + (ctx.device.id ? ' (' + ctx.device.id.slice(0, 8) + ')' : '');
}

/** 全角英数字→半角（normalize が使えない環境向けの保険つき） */
function nfkc_(s) {
  s = String(s == null ? '' : s);
  try { s = s.normalize('NFKC'); } catch (err) { /* 下の置換で対応 */ }
  return s.replace(/[\uFF01-\uFF5E]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
    .replace(/\u3000/g, ' ');
}

function normalizeSerial_(v) {
  return nfkc_(v).replace(/[\u0000-\u001F\u007F]/g, '').replace(/\s+/g, '').toUpperCase();
}

function normKey_(v) {
  return nfkc_(v).replace(/\s+/g, '').toUpperCase();
}

/** シートに書く文字列の安全化（= + - @ で始まる文字や、数字だけの文字列を「文字列」として保存する） */
function sheetText_(v) {
  if (typeof v !== 'string' || v === '') return v;
  if (/^[=+\-@]/.test(v) || /^[0-9.,\s]+$/.test(v)) return "'" + v;
  return v;
}

function cleanText_(v, max) {
  if (v == null) return '';
  return String(v).replace(/[\u0000-\u0008\u000B-\u001F\u007F]/g, '').trim().slice(0, max || 200);
}

function sanitizeClient_(d) {
  d = d || {};
  return { id: cleanText_(d.id, 64), label: cleanText_(d.label, 60) || '不明な端末' };
}

function toMs_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v.getTime();
  if (v === '' || v == null) return null;
  const n = new Date(v).getTime();
  return isNaN(n) ? null : n;
}

function fmtDate_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy/MM/dd HH:mm');
}

function newOpId_() {
  return 'OP' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + '-' + Utilities.getUuid().slice(0, 6);
}

function uniq_(a) {
  const seen = {};
  return a.filter(function (x) { if (seen[x]) return false; seen[x] = true; return true; });
}

function toSet_(a) {
  const s = {};
  a.forEach(function (x) { s[x] = true; });
  return s;
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

;
/**
 * ============================================================
 *  impactTV 端末管理 ― 管理用の機能（Admin.gs）
 *  初期設定・取込・台帳チェック・配布用QRコード
 *  ※ スプレッドシートのメニュー「impactTV管理」から実行します
 * ============================================================
 */

const DEFAULT_STATUSES = [
  // 状態, 種別, 色, 説明
  ['稼働中', '正常', '#15803d', '店舗で正常に使っている'],
  ['故障・不具合', '故障', '#dc2626', '映らない・電源が入らない など'],
  ['修理・交換中', '修理', '#c2410c', 'メーカーへ修理・交換を依頼中'],
  ['未使用・保管', '保管', '#1d4ed8', '使っていない／倉庫・本部で保管'],
  ['移動中', '移動', '#7c3aed', '別の場所へ運んでいる途中'],
  ['所在不明', '不明', '#334155', '棚卸で見つからなかった'],
  ['解約・返却', '終了', '#6b7280', '解約済み・返却済み（棚卸・長期未確認の対象外）'],
];
const DEFAULT_SPOTS = ['セット台', 'ネイル台', '受付', '待合', '店販棚', '大型・縦型モニター', 'その他'];
const DEFAULT_SYMPTOMS = ['電源が入らない', '画面が映らない・暗い', '画面の割れ・破損', '映像が止まる・乱れる', '内容が更新されない', 'Wi-Fiにつながらない', 'スタンド・金具の破損', 'その他'];
const DEFAULT_MODELS = [['PAC07', '7インチ'], ['PAS07', '7インチ'], ['PAC10', '10インチ'], ['PAS10', '10インチ'], ['IS700', 'STB（大型・縦型モニター用）']];
const DEFAULT_CONTRACTS = ['契約中', '解約申請中', '解約済'];
const KIND_HELP = [
  ['正常', '通常の状態。新規登録や「直った」のときに使います'],
  ['故障', '症状を選ぶ欄が出ます。集計の「要対応」に出ます'],
  ['修理', '店舗にない扱い（棚卸の対象外）。「要対応」に出ます'],
  ['保管', '使っていない端末（棚卸の対象）'],
  ['移動', '運んでいる途中（棚卸で見つかれば「稼働中」に戻せます）'],
  ['不明', '棚卸で見つからなかった端末。「要対応」に出ます'],
  ['終了', '解約・返却・廃棄（棚卸・長期未確認の対象外）'],
];
const OPTION_COLS = { status: 1, kind: 2, color: 3, desc: 4, spot: 6, symptom: 8, modelPrefix: 10, modelName: 11, contract: 13, kindList: 15, kindHelp: 16 };
const HEADER_BG = '#1f3a5f';

/* ============================================================
 *  初期設定
 * ============================================================ */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートの「拡張機能 → Apps Script」から開いたプロジェクトで実行してください。');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SS_ID', ss.getId());
  if (!props.getProperty('TOKEN_SECRET')) props.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  try { ss.setSpreadsheetTimeZone(TZ); } catch (err) { /* 何もしない */ }

  const dev = ensureSheet_(ss, SHEET.DEVICES, DEVICE_HEADERS, 1);
  const log = ensureSheet_(ss, SHEET.LOG, LOG_HEADERS, 2);
  const sum = ensureSheet_(ss, SHEET.SUMMARY, null, 3);
  const imp = ensureSheet_(ss, SHEET.IMPORT, null, 4);
  const st = ensureSheet_(ss, SHEET.STORES, STORE_HEADERS, 5);
  const op = ensureSheet_(ss, SHEET.OPTIONS, null, 6);
  const se = ensureSheet_(ss, SHEET.SETTINGS, ['項目', '値', '説明'], 7);

  setupOptions_(op);
  setupStores_(st);
  const passcode = setupSettings_(ss, se);
  SpreadsheetApp.flush();
  formatDevices_(ss, dev);
  formatLog_(log);
  setupImport_(imp);
  setupSummary_(ss, sum);
  removeDefaultSheet_(ss);
  CacheService.getScriptCache().remove(SETTINGS_CACHE_KEY);
  ss.setActiveSheet(dev);

  SpreadsheetApp.getUi().alert('初期設定が完了しました',
    '合言葉：' + passcode + '\n（「設定」シートで変更できます）\n\n' +
    '次の手順\n' +
    '① 「取込」シートに、各店サイネージ利用状況の「サイネージ使用状況」シートを見出しごと貼り付け\n' +
    '② メニュー「impactTV管理」→「取込」シートの内容を台帳に取り込む\n' +
    '③ 「デプロイ」→「新しいデプロイ」→ ウェブアプリ（実行：自分／アクセス：全員）\n' +
    '④ 発行されたURLを「設定」シートの「アプリURL」に貼り付け → メニューから配布用QRコードを表示',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function ensureSheet_(ss, name, headers, position) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, Math.min(position - 1, ss.getSheets().length));
  if (headers) {
    const lastCol = sh.getLastColumn();
    const current = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v).trim(); }) : [];
    if (!current.some(Boolean)) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      const missing = headers.filter(function (h) { return current.indexOf(h) < 0; });
      if (missing.length) sh.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
    }
    styleHeader_(sh);
  }
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(Math.min(position, ss.getSheets().length));
  return sh;
}

function styleHeader_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  sh.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground(HEADER_BG).setFontColor('#ffffff').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
}

function setupOptions_(sh) {
  const empty = !String(sh.getRange(1, OPTION_COLS.status).getValue()).trim();
  if (empty) {
    sh.clear();
    const put = function (col, header, list) {
      sh.getRange(1, col).setValue(header);
      if (list.length) sh.getRange(2, col, list.length, 1).setValues(list.map(function (v) { return [v]; }));
    };
    sh.getRange(1, OPTION_COLS.status, 1, 4).setValues([['状態', '種別', '色', '説明']]);
    sh.getRange(2, OPTION_COLS.status, DEFAULT_STATUSES.length, 4).setValues(DEFAULT_STATUSES);
    put(OPTION_COLS.spot, '設置個所', DEFAULT_SPOTS);
    put(OPTION_COLS.symptom, '症状', DEFAULT_SYMPTOMS);
    sh.getRange(1, OPTION_COLS.modelPrefix, 1, 2).setValues([['機種の先頭文字', '機種名']]);
    sh.getRange(2, OPTION_COLS.modelPrefix, DEFAULT_MODELS.length, 2).setValues(DEFAULT_MODELS);
    put(OPTION_COLS.contract, '契約', DEFAULT_CONTRACTS);
  }
  // 種別の説明表（毎回書き直す）
  sh.getRange(1, OPTION_COLS.kindList, KIND_HELP.length + 1, 2).setValues([['種別の一覧', '意味']].concat(KIND_HELP));
  sh.getRange(1, 1, 1, OPTION_COLS.kindHelp).setFontWeight('bold').setBackground(HEADER_BG).setFontColor('#ffffff');
  [OPTION_COLS.desc + 1, OPTION_COLS.spot + 1, OPTION_COLS.symptom + 1, OPTION_COLS.modelName + 1, OPTION_COLS.contract + 1].forEach(function (c) {
    sh.getRange(1, c).setBackground(null).setValue('');
    sh.setColumnWidth(c, 24);
  });
  sh.setFrozenRows(1);
  sh.setColumnWidth(OPTION_COLS.status, 120);
  sh.setColumnWidth(OPTION_COLS.desc, 260);
  sh.setColumnWidth(OPTION_COLS.spot, 130);
  sh.setColumnWidth(OPTION_COLS.symptom, 170);
  sh.setColumnWidth(OPTION_COLS.modelPrefix, 110);
  sh.setColumnWidth(OPTION_COLS.modelName, 200);
  sh.setColumnWidth(OPTION_COLS.kindHelp, 380);
  // 種別はプルダウン
  const kindRule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_KINDS, true).setAllowInvalid(false).build();
  sh.getRange(2, OPTION_COLS.kind, 48, 1).setDataValidation(kindRule);
  // 色のセルをその色で塗る
  const last = sh.getLastRow();
  if (last >= 2) {
    const colors = sh.getRange(2, OPTION_COLS.color, last - 1, 1).getValues();
    colors.forEach(function (r, i) {
      const c = String(r[0]).trim();
      if (/^#[0-9a-f]{6}$/i.test(c)) sh.getRange(i + 2, OPTION_COLS.status, 1, 1).setBackground(c).setFontColor('#ffffff').setFontWeight('bold');
    });
  }
  sh.getRange(1, OPTION_COLS.status).setNote('状態は行を追加・名前変更できます。\n「種別」は必ず選んでください（アプリの動きが決まります）。\n色は #15803d のような形式です。\n変更後はメニュー「初期設定」を再実行すると台帳の色分けも更新されます。');
}

function setupStores_(sh) {
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 1, STORE_HEADERS.length).setValues([['本部在庫', '本部', 0, false, '予備・返却品の保管場所']]);
  }
  const hideCol = STORE_HEADERS.indexOf('非表示') + 1;
  sh.getRange(2, hideCol, 998, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 80);
  sh.setColumnWidth(3, 70);
  sh.setColumnWidth(4, 70);
  sh.setColumnWidth(5, 260);
  sh.getRange(1, 1).setNote('アプリの「設置場所」に出る一覧です。\n表示順：小さい順に上に表示（空欄は下に行の順で表示）\n非表示：チェックすると選択肢に出なくなります（台帳の値はそのまま）');
}

function setupSettings_(ss, sh) {
  const values = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), Math.max(sh.getLastColumn(), 3)).getValues() : [];
  const rowOf = {};
  values.forEach(function (r, i) { rowOf[String(r[0]).trim()] = i + 1; });
  SETTING_DEFS.forEach(function (d) {
    if (!rowOf[d[1]]) {
      const r = sh.getLastRow() + 1;
      sh.getRange(r, 1, 1, 3).setValues([[d[1], d[2], d[3]]]);
      rowOf[d[1]] = r;
    }
  });
  // 合言葉
  const passRow = rowOf['合言葉'];
  let passcode = String(sh.getRange(passRow, 2).getValue()).trim();
  if (!passcode) {
    const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
    for (let i = 0; i < 6; i++) passcode += chars.charAt(Math.floor(Math.random() * chars.length));
    sh.getRange(passRow, 2).setValue(passcode);
  }
  sh.getRange(passRow, 2).setNumberFormat('@').setFontWeight('bold').setBackground('#fef9c3');
  // 写真フォルダ
  const folderRow = rowOf['写真フォルダID'];
  const folderId = String(sh.getRange(folderRow, 2).getValue()).trim();
  let ok = false;
  if (folderId) { try { DriveApp.getFolderById(folderId); ok = true; } catch (err) { ok = false; } }
  if (!ok) {
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    sh.getRange(folderRow, 2).setValue(parent.createFolder('impactTV端末管理_写真').getId());
  }
  // 名前付き範囲（集計で使う）
  const staleCell = sh.getRange(rowOf['長期未確認（日）'], 2);
  const named = ss.getNamedRanges().filter(function (n) { return n.getName() === 'STALE_DAYS'; })[0];
  if (named) named.setRange(staleCell); else ss.setNamedRange('STALE_DAYS', staleCell);
  sh.getRange(rowOf['シリアル形式'], 2).setNumberFormat('@');
  styleHeader_(sh);
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(2, 300);
  sh.setColumnWidth(3, 520);
  return passcode;
}

function formatDevices_(ss, sh) {
  const t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
  const col = function (h) { return t.idx[h] + 1; };
  const letter = function (h) { return colLetter_(col(h)); };
  const maxRows = Math.max(sh.getMaxRows(), 1000);
  if (sh.getMaxRows() < 1000) sh.insertRowsAfter(sh.getMaxRows(), 1000 - sh.getMaxRows());
  sh.setFrozenColumns(1);
  const widths = { シリアル番号: 150, 設置場所: 190, 設置個所: 110, 状態: 110, 症状: 200, 機種: 150, 契約: 90, メモ: 260, 最終確認日時: 130, 最終確認者: 90, 更新日時: 130, 更新者: 90, 写真: 120, 登録日時: 130, 版: 45, 最終操作ID: 150 };
  Object.keys(widths).forEach(function (h) { if (h in t.idx) sh.setColumnWidth(col(h), widths[h]); });
  [D.checkedAt, D.updatedAt, D.createdAt].forEach(function (h) {
    sh.getRange(2, col(h), maxRows - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  });
  sh.getRange(2, col(D.serial), maxRows - 1, 1).setNumberFormat('@');
  [D.rev, D.lastOp].forEach(function (h) { sh.getRange(1, col(h)).setBackground('#64748b'); });
  sh.getRange(1, col(D.rev)).setNote('アプリが使う管理用の列です（編集しないでください）');

  // プルダウン（候補外の値は警告のみ）
  const optSheet = ss.getSheetByName(SHEET.OPTIONS);
  const storeSheet = ss.getSheetByName(SHEET.STORES);
  const listRule = function (range) {
    return SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(true).build();
  };
  sh.getRange(2, col(D.location), maxRows - 1, 1).setDataValidation(listRule(storeSheet.getRange('A2:A1000')));
  sh.getRange(2, col(D.status), maxRows - 1, 1).setDataValidation(listRule(optSheet.getRange(2, OPTION_COLS.status, 49, 1)));
  sh.getRange(2, col(D.spot), maxRows - 1, 1).setDataValidation(listRule(optSheet.getRange(2, OPTION_COLS.spot, 49, 1)));
  sh.getRange(2, col(D.contract), maxRows - 1, 1).setDataValidation(listRule(optSheet.getRange(2, OPTION_COLS.contract, 49, 1)));

  // 条件付き書式（重複・状態の色・長期未確認）
  const rules = [];
  const A = letter(D.serial);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + A + '2<>"",COUNTIF($' + A + ':$' + A + ',$' + A + '2)>1)')
    .setBackground('#fecaca').setFontColor('#991b1b').setBold(true)
    .setRanges([sh.getRange(A + '2:' + A)]).build());
  const statusRange = sh.getRange(letter(D.status) + '2:' + letter(D.status));
  readOptions_(optSheet).statuses.forEach(function (s) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s.name)
      .setBackground(tint_(s.color)).setFontColor(s.color).setBold(true).setRanges([statusRange]).build());
  });
  const I = letter(D.checkedAt);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($' + A + '2<>"",$' + I + '2<TODAY()-INDIRECT("STALE_DAYS"))')
    .setBackground('#fef3c7').setRanges([sh.getRange(I + '2:' + I)]).build());
  sh.setConditionalFormatRules(rules);
  sh.getRange(1, col(D.serial)).setNote('赤色＝同じシリアルが複数行あります（重複）\n最終確認日時が黄色＝長期間読み取られていない端末');
}

function formatLog_(sh) {
  const t = openTable_(SHEET.LOG, LOG_HEADERS);
  const col = function (h) { return t.idx[h] + 1; };
  const widths = { 日時: 130, 操作ID: 150, シリアル番号: 140, 操作: 120, '状態(前)': 100, '状態(後)': 100, '場所(前)': 160, '場所(後)': 160, 設置個所: 90, 症状: 180, コメント: 220, 写真: 120, 担当者: 90, 端末: 160, 取消: 160 };
  Object.keys(widths).forEach(function (h) { if (h in t.idx) sh.setColumnWidth(col(h), widths[h]); });
  sh.getRange(2, col(L.time), Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  sh.hideColumns(col(L.snap));
  sh.getRange(1, col(L.time)).setNote('アプリの操作がすべて記録されます（編集・削除しないでください）');
}

function setupImport_(sh) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, IMPORT_HEADERS.length).setValues([IMPORT_HEADERS]);
    styleHeader_(sh);
  }
  sh.getRange(1, 1).setNote('ここに「各店サイネージ利用状況」の「サイネージ使用状況」シートを、見出し行ごと貼り付けてください。\n' +
    '見出しに「端末番号」または「シリアル」を含む列がシリアルとして使われます。\n' +
    '貼り付けたら、メニュー「impactTV管理」→「「取込」シートの内容を台帳に取り込む」を実行します。\n' +
    '台帳にすでにあるシリアルはスキップされるので、何度実行しても重複しません。');
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 190);
  sh.setColumnWidth(4, 300);
}

function setupSummary_(ss, sh) {
  const t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
  const L_ = function (h) { return colLetter_(t.idx[h] + 1); };
  const lastLetter = colLetter_(t.width);
  const lg = openTable_(SHEET.LOG, LOG_HEADERS);
  const LL = function (h) { return colLetter_(lg.idx[h] + 1); };
  const dev = "'" + SHEET.DEVICES + "'!";
  const opt = "'" + SHEET.OPTIONS + "'!";
  const logR = "'" + SHEET.LOG + "'!";
  const col = function (h) { return dev + L_(h) + '2:' + L_(h); };
  const kinds = function (list) {
    return 'FILTER(' + opt + 'A2:A,' + list.map(function (k) { return '(' + opt + 'B2:B="' + k + '")'; }).join('+') + ')';
  };
  sh.clear();
  const blocks = [
    ['A1', '状態別の台数',
      '=QUERY(' + dev + 'A:' + lastLetter + ',"select ' + L_(D.status) + ', count(' + L_(D.serial) + ') where ' + L_(D.serial) + " <> '' group by " + L_(D.status) + " label " + L_(D.status) + " '状態', count(" + L_(D.serial) + ") '台数'\",1)"],
    ['D1', '設置場所 × 状態',
      '=QUERY(' + dev + 'A:' + lastLetter + ',"select ' + L_(D.location) + ', count(' + L_(D.serial) + ') where ' + L_(D.serial) + " <> '' group by " + L_(D.location) + ' pivot ' + L_(D.status) + ' label ' + L_(D.location) + " '設置場所'\",1)"],
    ['P1', '要対応（故障・修理・所在不明）',
      '=IFERROR(SORT(FILTER({' + [D.serial, D.location, D.status, D.symptom, D.updatedAt].map(col).join(',') + '},ISNUMBER(MATCH(' + col(D.status) + ',' + kinds(['故障', '修理', '不明']) + ',0))),2,TRUE),"該当なし")'],
    ['V1', '未確認・長期未確認（設定の日数以上）',
      '=IFERROR(SORT(FILTER({' + [D.serial, D.location, D.status, D.checkedAt].map(col).join(',') + '},' + col(D.serial) + '<>"",' + col(D.checkedAt) + '<TODAY()-STALE_DAYS,ISNA(MATCH(' + col(D.status) + ',' + kinds(['終了']) + ',0))),2,TRUE,1,TRUE),"該当なし")'],
    ['AA1', '重複しているシリアル',
      '=IFERROR(UNIQUE(FILTER(' + col(D.serial) + ',' + col(D.serial) + '<>"",COUNTIF(' + col(D.serial) + ',' + col(D.serial) + ')>1)),"重複なし")'],
    ['AC1', '最近の操作（30件）',
      '=QUERY(' + logR + 'A:' + colLetter_(lg.width) + ',"select ' + [L.time, L.serial, L.op, L.st1, L.l1, L.user].map(LL).join(', ') + ' where ' + LL(L.time) + ' is not null order by ' + LL(L.time) + ' desc limit 30",1)'],
  ];
  const heads = {
    P1: ['シリアル番号', '設置場所', '状態', '症状', '更新日時'],
    V1: ['シリアル番号', '設置場所', '状態', '最終確認日時'],
  };
  blocks.forEach(function (b) {
    const cell = sh.getRange(b[0]);
    cell.setValue(b[1]).setFontWeight('bold').setFontSize(12);
    const r = cell.getRow();
    const c = cell.getColumn();
    if (heads[b[0]]) {
      sh.getRange(r + 1, c, 1, heads[b[0]].length).setValues([heads[b[0]]]).setFontWeight('bold').setBackground('#e2e8f0');
      sh.getRange(r + 2, c).setFormula(b[2]);
    } else {
      sh.getRange(r + 1, c).setFormula(b[2]);
    }
  });
  sh.getRange('B1').setFormula('="全 "&COUNTIF(' + col(D.serial) + ',"?*")&" 台"').setFontColor('#475569');
  sh.getRange('T:T').setNumberFormat('yyyy/mm/dd hh:mm');
  sh.getRange('Y:Y').setNumberFormat('yyyy/mm/dd');
  sh.getRange('AC:AC').setNumberFormat('yyyy/mm/dd hh:mm');
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(4, 190);
  [16, 22, 27, 29].forEach(function (c) { sh.setColumnWidth(c, 140); });
  sh.setColumnWidth(17, 170);
  sh.setColumnWidth(23, 170);
  sh.setColumnWidth(31, 110);
  sh.setColumnWidth(33, 170);
  sh.setTabColor('#0f766e');
}

function removeDefaultSheet_(ss) {
  ['シート1', 'Sheet1'].forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}

function tint_(hex) {
  const n = parseInt(hex.slice(1), 16);
  const mix = function (c) { return Math.round(c + (255 - c) * 0.85); };
  const r = mix((n >> 16) & 255), g = mix((n >> 8) & 255), b = mix(n & 255);
  return '#' + [r, g, b].map(function (x) { return ('0' + x.toString(16)).slice(-2); }).join('');
}

/* ============================================================
 *  取込（PISTAの端末台帳など）
 * ============================================================ */

function importFromSheet() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('台帳への取込', '「取込」シートの内容を台帳に追加します。\n台帳にすでにあるシリアルはスキップします。\n\n実行しますか？', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  const res = withLock_(function () { return runImport_(); });
  ui.alert('取込結果', importSummaryText_(res), ui.ButtonSet.OK);
}

function importFromSpreadsheetUrl() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.prompt('別のスプレッドシートから取り込む',
    'GoogleスプレッドシートのURLを貼り付けてください。\n（Excelファイルは、先に「Googleスプレッドシートとして保存」してください）',
    ui.ButtonSet.OK_CANCEL);
  if (ans.getSelectedButton() !== ui.Button.OK) return;
  const url = ans.getResponseText().trim();
  let src;
  try { src = SpreadsheetApp.openByUrl(url); } catch (err) {
    ui.alert('開けませんでした', 'URLが正しいか、このアカウントで閲覧できるファイルか確認してください。\n' + err.message, ui.ButtonSet.OK);
    return;
  }
  let found = null;
  src.getSheets().some(function (s) {
    const lastRow = s.getLastRow(), lastCol = s.getLastColumn();
    if (!lastRow || !lastCol) return false;
    const values = s.getRange(1, 1, lastRow, lastCol).getValues();
    for (let r = 0; r < Math.min(values.length, 10); r++) {
      if (values[r].some(function (v) { return /端末番号|シリアル/.test(String(v)); })) { found = { sheet: s, values: values.slice(r) }; return true; }
    }
    return false;
  });
  if (!found) {
    ui.alert('見つかりませんでした', '「端末番号」または「シリアル」という見出しのあるシートがありませんでした。', ui.ButtonSet.OK);
    return;
  }
  const ok = ui.alert('確認', '「' + src.getName() + '」の「' + found.sheet.getName() + '」（' + (found.values.length - 1) +
    '行）を「取込」シートに写して、台帳に取り込みます。\n「取込」シートの今の内容は消えます。よろしいですか？', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  const imp = getSheet_(SHEET.IMPORT);
  imp.clearContents();
  const width = Math.max.apply(null, found.values.map(function (r) { return r.length; }));
  const safe = found.values.map(function (r) {
    const line = r.map(function (v) { return typeof v === 'string' ? sheetText_(v) : v; });
    while (line.length < width) line.push('');
    return line;
  });
  imp.getRange(1, 1, safe.length, width).setValues(safe);
  const res = withLock_(function () { return runImport_(); });
  ui.alert('取込結果', importSummaryText_(res), ui.ButtonSet.OK);
}

/** 取込の本体（画面を使わないのでテストしやすい形） */
function runImport_() {
  const S = loadSettings_();
  const sh = getSheet_(SHEET.IMPORT);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) throw new Error('「取込」シートにデータがありません。');
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  let hr = -1;
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    if (values[r].some(function (v) { return /端末番号|シリアル/.test(String(v)); })) { hr = r; break; }
  }
  if (hr < 0) throw new Error('見出し行（「端末番号」または「シリアル」を含む行）が見つかりません。');
  const H = values[hr].map(function (v) { return String(v).trim(); });
  const find = function (re, not) {
    for (let i = 0; i < H.length; i++) if (re.test(H[i]) && !(not && not.test(H[i]))) return i;
    return -1;
  };
  const c = {
    serial: find(/端末番号|シリアル/),
    store: find(/店舗|設置場所/, /個所|箇所/),
    spot: find(/設置個所|設置箇所/, /その他|どこ/),
    detail: find(/その他|どこに/),
    cancel: find(/^解約/),
    message: find(/メッセージ|備考/),
    status: find(/^状態/),
    model: find(/機種/),
    contract: find(/契約/),
    memo: find(/^メモ/),
    result: find(/取込結果/),
  };
  if (c.store < 0) throw new Error('「店舗名」（または「設置場所」）の列が見つかりません。');
  if (c.result < 0) {
    c.result = H.length;
    sh.getRange(hr + 1, c.result + 1).setValue('取込結果').setFontWeight('bold');
  }

  const t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
  const existing = indexDevices_(t);
  const stores = S.stores.slice();
  const storeKey = {};
  stores.forEach(function (s) { storeKey[normKey_(s.name)] = s.name; });
  const newStores = [];
  const ensureStore = function (name) {
    const k = normKey_(name);
    if (!storeKey[k]) { storeKey[k] = name; newStores.push(name); }
    return storeKey[k];
  };
  const pattern = S.serialPattern ? new RegExp(S.serialPattern) : null;
  const contractDefault = S.contracts[0] || '';
  const contractEnded = S.contracts.filter(function (x) { return /解約済|解約$/.test(x); })[0] || '解約済';
  const now = new Date();
  const opId = newOpId_();
  const seen = {};
  const devices = [];
  const logs = [];
  const results = [];
  const res = { added: 0, exists: 0, dup: 0, skipped: 0, review: 0, badFormat: 0, byStatus: {}, newStores: newStores, reviewList: [] };

  for (let r = hr + 1; r < values.length; r++) {
    const row = values[r];
    const cell = function (i) { return i >= 0 ? cleanText_(row[i], 300) : ''; };
    const serial = normalizeSerial_(row[c.serial]);
    let result = '';
    if (!serial) {
      result = row.some(function (v, i) { return i !== c.result && String(v).trim(); }) ? 'シリアルなし（スキップ）' : '';
      if (result) res.skipped++;
    } else if (!/^[A-Z0-9\-_.\/+]{1,40}$/.test(serial)) {
      result = 'シリアルの形式が正しくありません（スキップ）';
      res.skipped++;
    } else if (seen[serial]) {
      result = '重複：' + seen[serial] + '行目と同じシリアル（スキップ）';
      res.dup++;
    } else if (existing[serial]) {
      seen[serial] = r + 1;
      result = '台帳に登録済み（スキップ）';
      res.exists++;
    } else if (!cell(c.store)) {
      seen[serial] = r + 1;
      result = '店舗名がありません（スキップ）';
      res.skipped++;
    } else {
      seen[serial] = r + 1;
      const store = cell(c.store);
      const detail = cell(c.detail);
      const message = cell(c.message);
      const cancel = c.cancel >= 0 && isMark_(row[c.cancel]);
      const statusText = cell(c.status);
      const given = statusText && S.statuses.some(function (x) { return x.name === statusText; }) ? statusText : null;
      const guess = inferImportStatus_(S, cancel, [detail, message, given ? '' : statusText].filter(Boolean).join(' '));
      const status = given || guess.status;
      const origin = ensureStore(store);
      const location = guess.toHome ? ensureStore(S.homeLocation) : origin;
      const brokenText = statusKind_(S, status) === '故障' ? (detail || message) : '';
      const memo = [cell(c.memo), brokenText === detail ? '' : detail, brokenText === message ? '' : message, guess.toHome ? '元の設置：' + store : ''].filter(Boolean).join(' / ');
      let contract = cell(c.contract);
      if (cancel) contract = contractEnded;
      if (!contract && !given && !statusText) contract = contractDefault;
      const dev = {
        serial: serial, location: location, spot: cell(c.spot), status: status, symptom: brokenText.slice(0, 200),
        model: cell(c.model) || inferModel_(S, serial), contract: contract, memo: memo.slice(0, 500),
        checkedAt: '', checkedBy: '', updatedAt: now, updatedBy: '本部・取込', photo: '', createdAt: now, rev: 1, lastOp: '',
      };
      devices.push(dev);
      logs.push({
        time: now, opId: opId, serial: serial, op: '台帳に取込', st0: '', st1: status, l0: '', l1: location,
        spot: dev.spot, symptom: dev.symptom, memo: [detail, message].filter(Boolean).join(' / ').slice(0, 300), photo: '', user: '本部', device: 'スプレッドシート',
        undone: '', snap: '',
      });
      res.added++;
      res.byStatus[status] = (res.byStatus[status] || 0) + 1;
      result = '取込OK（' + status + (guess.toHome ? '・' + S.homeLocation + 'へ' : '') + '）';
      if (pattern && !pattern.test(serial)) { result += '／形式が違います'; res.badFormat++; }
      if (guess.review) {
        result += '／要確認：' + (detail || message).slice(0, 40);
        res.review++;
        res.reviewList.push(serial + '（' + store + '）→ ' + status);
      }
    }
    results.push([result]);
  }

  if (newStores.length) {
    const st = getSheet_(SHEET.STORES);
    const stamp = Utilities.formatDate(now, TZ, 'yyyy/MM/dd');
    st.getRange(st.getLastRow() + 1, 1, newStores.length, STORE_HEADERS.length).setValues(newStores.map(function (n) {
      return [sheetText_(n), n === S.homeLocation ? '本部' : '店舗', '', false, '取込で追加 ' + stamp];
    }));
  }
  appendDevices_(t, devices);
  appendLogs_(logs);
  if (results.length) sh.getRange(hr + 2, c.result + 1, results.length, 1).setValues(results);
  CacheService.getScriptCache().remove(SETTINGS_CACHE_KEY);
  return res;
}

/** 取込時の状態の推定（「→」があれば、その後ろ＝今の状態として読む） */
function inferImportStatus_(S, cancel, text) {
  const K = function (kind) { return kindStatus_(S, kind); };
  const whole = String(text || '');
  const parts = whole.split('→');
  const tail = parts[parts.length - 1];
  const review = !cancel && (/故障|不具合|返却|返品|回収|未使用|在庫|保管|解約|交換|紛失|不明/.test(tail) || /確認中|？|\?/.test(whole));
  let status = null;
  let toHome = false;
  if (cancel) status = K('終了');
  else if (/返却|返品/.test(tail)) {
    if (/本部/.test(tail)) { status = K('保管'); toHome = true; } else status = K('終了');
  } else if (/本部在庫/.test(tail) || (/回収/.test(tail) && /本部|在庫/.test(tail))) { status = K('保管'); toHome = true; }
  else if (/故障|不具合|破損|割れ|ﾜﾚ|ワレ/.test(tail)) status = K('故障');
  else if (/未使用|在庫|保管|使用なし/.test(tail)) status = K('保管');
  else if (/紛失|所在不明/.test(tail)) status = K('不明');
  return { status: status || defaultStatus_(S), toHome: toHome && !!S.homeLocation, review: review };
}

function isMark_(v) {
  if (v === true) return true;
  return /^(〇|○|◯|●|✓|✔|済|1|TRUE|はい|YES|解約)/i.test(String(v == null ? '' : v).trim());
}

function importSummaryText_(res) {
  const lines = [];
  lines.push('台帳に追加：' + res.added + '件');
  Object.keys(res.byStatus).forEach(function (k) { lines.push('　・' + k + '：' + res.byStatus[k] + '件'); });
  lines.push('登録済みでスキップ：' + res.exists + '件');
  if (res.dup) lines.push('取込データ内の重複：' + res.dup + '件');
  if (res.skipped) lines.push('その他スキップ：' + res.skipped + '件');
  if (res.badFormat) lines.push('シリアル形式が違う：' + res.badFormat + '件');
  if (res.newStores.length) lines.push('店舗マスタに追加：' + res.newStores.length + '件（' + res.newStores.slice(0, 8).join('、') + (res.newStores.length > 8 ? ' ほか' : '') + '）');
  if (res.review) {
    lines.push('');
    lines.push('要確認（メモから状態を推定）：' + res.review + '件');
    res.reviewList.slice(0, 12).forEach(function (x) { lines.push('　' + x); });
    if (res.reviewList.length > 12) lines.push('　…「取込」シートの取込結果列を確認してください');
  }
  return lines.join('\n');
}

/* ============================================================
 *  台帳のチェック
 * ============================================================ */

function checkLedger() {
  const r = runLedgerCheck_();
  const lines = [];
  const add = function (title, list) {
    lines.push(title + '：' + list.length + '件');
    list.slice(0, 10).forEach(function (x) { lines.push('　' + x); });
    if (list.length > 10) lines.push('　…ほか' + (list.length - 10) + '件');
  };
  add('重複しているシリアル', r.duplicates);
  add('シリアルが空の行', r.blanks);
  add('設置場所が店舗マスタにない', r.unknownLocations);
  add('状態が「選択肢」にない', r.unknownStatuses);
  add('シリアル形式が違う', r.badFormat);
  SpreadsheetApp.getUi().alert('台帳のチェック（全' + r.total + '台）', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function runLedgerCheck_() {
  const S = loadSettings_();
  const t = readTable_(SHEET.DEVICES, DEVICE_HEADERS);
  const index = {};
  const res = { total: 0, duplicates: [], blanks: [], unknownLocations: [], unknownStatuses: [], badFormat: [] };
  const storeKeys = {};
  S.stores.forEach(function (s) { storeKeys[normKey_(s.name)] = true; });
  const statusNames = S.statuses.map(function (s) { return s.name; });
  const pattern = S.serialPattern ? new RegExp(S.serialPattern) : null;
  t.rows.forEach(function (row, i) {
    const d = rowToDevice_(t, row, i + 2);
    const hasData = row.some(function (v) { return String(v).trim(); });
    if (!d.serial) { if (hasData) res.blanks.push((i + 2) + '行目'); return; }
    res.total++;
    (index[d.serial] = index[d.serial] || []).push(i + 2);
    if (d.location && !storeKeys[normKey_(d.location)]) res.unknownLocations.push((i + 2) + '行目 ' + d.serial + '：' + d.location);
    if (statusNames.indexOf(d.status) < 0) res.unknownStatuses.push((i + 2) + '行目 ' + d.serial + '：' + (d.status || '（空欄）'));
    if (pattern && !pattern.test(d.serial)) res.badFormat.push((i + 2) + '行目 ' + d.serial);
  });
  Object.keys(index).forEach(function (s) {
    if (index[s].length > 1) res.duplicates.push(s + '（' + index[s].join('・') + '行目）');
  });
  return res;
}

/* ============================================================
 *  配布用URL・QRコード
 * ============================================================ */

function showShareDialog() {
  const S = loadSettings_();
  let url = S.appUrl;
  if (!url) { try { url = ScriptApp.getService().getUrl() || ''; } catch (err) { url = ''; } }
  const html = HtmlService.createHtmlOutput(buildShareHtml_(url, S.passcode)).setWidth(460).setHeight(660);
  SpreadsheetApp.getUi().showModalDialog(html, '配布用URL・QRコード');
}

function buildShareHtml_(url, passcode) {
  const sep = /script\.google\.com/.test(url) ? (url.indexOf('?') >= 0 ? '&' : '?') : '#';
  const withKey = url ? url + sep + 'key=' + encodeURIComponent(passcode) : '';
  const esc = function (s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  if (!url) {
    return '<div style="font-family:sans-serif;line-height:1.7">「設定」シートの「アプリURL」が空です。<br>' +
      'デプロイ後に表示される「ウェブアプリのURL（…/exec）」を貼り付けてから、もう一度開いてください。</div>';
  }
  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<script src="https://cdn.jsdelivr.net/npm/qrcode-generator@1.4.4/qrcode.min.js"></script>' +
    '<style>body{font-family:system-ui,"Hiragino Sans","Noto Sans JP",sans-serif;margin:0;padding:8px 4px;color:#0f172a}' +
    'h3{font-size:15px;margin:14px 0 6px}.qr{display:flex;gap:12px;align-items:center}.qr div{flex:0 0 auto}' +
    'input{width:100%;font-size:12px;padding:6px;box-sizing:border-box}p{font-size:12px;color:#475569;margin:4px 0}</style></head><body>' +
    '<h3>① 合言葉つき（読み取るだけで使い始められます）</h3>' +
    '<div class="qr"><div id="q2"></div><p>社内の掲示・LINE WORKSでの共有用。<br>合言葉：<b>' + esc(passcode) + '</b><br>合言葉を変えたら作り直してください。</p></div>' +
    '<input readonly value="' + esc(withKey) + '" onclick="this.select()">' +
    '<h3>② URLのみ（開いたときに合言葉を入力）</h3>' +
    '<div class="qr"><div id="q1"></div><p>外部の人の目に触れる可能性がある場所にはこちら。</p></div>' +
    '<input readonly value="' + esc(url) + '" onclick="this.select()">' +
    '<script>function draw(id,text){try{var q=qrcode(0,"M");q.addData(text);q.make();document.getElementById(id).innerHTML=q.createSvgTag({cellSize:4,margin:2});}catch(e){document.getElementById(id).textContent="QRコードを作れませんでした";}}' +
    'draw("q1",' + JSON.stringify(url) + ');draw("q2",' + JSON.stringify(withKey) + ');</script>' +
    '</body></html>';
}

;return {api: api, setup: setup, runImport_: runImport_, runLedgerCheck_: runLedgerCheck_, loadSettings_: loadSettings_, makeToken_: makeToken_, onEdit: onEdit};
  })(G.SpreadsheetApp, G.DriveApp, G.Utilities, G.CacheService, G.PropertiesService, G.LockService, G.HtmlService, G.ContentService, G.ScriptApp, G.Session, G.console);
};


window.__LEDGER__ = [["端末番号(黄色シール)", "店舗名", "設置個所", "設置個所”その他”の場合どこに置いているか入力", "解約", "メッセージ"], ["PAC07A0900001", "デモ店A（本町）", "セット台", "", "", ""], ["PAC07A0900002", "デモ店A（本町）", "セット台", "", "", ""], ["PAC07A0900003", "デモ店A（本町）", "ネイル台", "", "", ""], ["PAC10A0900011", "デモ店A（本町）", "受付", "", "", ""], ["PAC07A0900021", "デモ店B（駅前）", "セット台", "", "", ""], ["PAC07A0900022", "デモ店B（駅前）", "セット台", "", "〇", "解約(故障)"], ["PAC10A0900031", "デモ店B（駅前）", "セット台", "", "", ""], ["IS700C0900041", "デモ店B（駅前）", "その他", "大型モニター", "", ""], ["PAC07A0900051", "デモ店C（港北）", "セット台", "", "", ""], ["PAC07A0900052", "デモ店C（港北）", "店販棚", "", "", "画面が映らない"], ["PAC07A0900061", "デモ店D（中央）", "セット台", "", "", ""], ["PAC07A0900071", "デモ店E（西口）", "セット台", "", "", "本部に返却"], ["PAC07A0900072", "デモ店E（西口）", "ネイル台", "", "", "本部に返却"], ["PAC10A0900081", "本部在庫", "セット台", "", "", ""]];

window.__startDemo__ = function () {
  var opts = {"passcode":"kaino","delay":250,"fakeGas":false,"preset":"お試し"};
  var env = GasMock.createGasEnv({ console: { error: function () {}, log: function () {} } });
  var server = window.__makeServer(env);
  server.setup();
  var ss = env.activeSS;
  var se = ss.getSheetByName('設定');
  var vals = se.getDataRange().getValues();
  for (var i = 0; i < vals.length; i++) if (vals[i][0] === '合言葉') se.getRange(i + 1, 2).setValue(opts.passcode);
  var L = window.__LEDGER__;
  var imp = ss.getSheetByName('取込');
  imp.clearContents();
  imp.getRange(1, 1, L.length, L[0].length).setValues(L);
  server.runImport_();
  Object.keys(env.cacheStore).forEach(function (k) { delete env.cacheStore[k]; });
  window.__SERVER__ = server;
  window.__ENV__ = env;
  var delay = opts.delay;
  function call(req) { return JSON.parse(JSON.stringify(server.api(JSON.parse(JSON.stringify(req))))); }
  window.__CALL__ = call;
  if (opts.fakeGas) {
    var Runner = function (s, f) { this._s = s; this._f = f; };
    Runner.prototype.withSuccessHandler = function (fn) { return new Runner(fn, this._f); };
    Runner.prototype.withFailureHandler = function (fn) { return new Runner(this._s, fn); };
    Runner.prototype.api = function (req) {
      var self = this;
      window.__CALLS__ = (window.__CALLS__ || 0) + 1;
      setTimeout(function () {
        if (window.__OFFLINE__) { if (self._f) self._f(new Error('NetworkError: Connection failure due to HTTP 0')); return; }
        var r;
        try { r = call(req); } catch (e) { if (self._f) self._f(e); return; }
        if (self._s) self._s(r);
      }, delay);
    };
    window.google = { script: {
      run: new Runner(null, null),
      url: { getLocation: function (cb) {
        var p = {}; new URLSearchParams(location.search).forEach(function (v, k) { p[k] = v; });
        setTimeout(function () { cb({ parameter: p, parameters: {}, hash: '' }); }, 5);
      } },
      history: {
        push: function (s, p, h) { window.__HIST__ = (window.__HIST__ || []).concat([h]); },
        replace: function () { window.__HIST_REPLACED__ = true; },
        setChangeHandler: function (fn) { window.__HIST_HANDLER__ = fn; }
      }
    } };
  } else {
    window.__MOCK_API__ = function (req) {
      window.__CALLS__ = (window.__CALLS__ || 0) + 1;
      return new Promise(function (res, rej) {
        setTimeout(function () { if (window.__OFFLINE__) rej(new Error('offline')); else res(call(req)); }, delay);
      });
    };
  }
  if (opts.config) window.ITV_CONFIG = opts.config;
  window.__DEMO__ = true;
  return Promise.resolve();
};
