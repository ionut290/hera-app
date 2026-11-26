import test from 'node:test';import assert from 'node:assert/strict';import {inspectXlsxZip,verifyXlsxInflation} from './portal/src/price-xlsx.mjs';
test('rejects invalid and oversized XLSX before ExcelJS parsing',()=>{assert.throws(()=>inspectXlsxZip(new ArrayBuffer(0)));assert.throws(()=>inspectXlsxZip(new ArrayBuffer(5000001)));assert.throws(()=>inspectXlsxZip(new ArrayBuffer(22)));});

import {deflateRawSync} from 'node:zlib';
test('actual ZIP inflation rejects forged sizes, accepts genuine content',async()=>{
 const bytes=deflateRawSync(Buffer.alloc(10000,65));const buffer=Uint8Array.from(bytes).buffer;
 const entry={dataAt:0,packed:bytes.length,size:10000,method:8};
 await verifyXlsxInflation(buffer,[entry]);
 await assert.rejects(verifyXlsxInflation(buffer,[{...entry,size:1}]),/limite/);
 await assert.rejects(verifyXlsxInflation(buffer,[{...entry,size:10001}]),/incoerenti/);
 await assert.rejects(verifyXlsxInflation(buffer,[{...entry,method:0}]),/incoerenti/);
});
