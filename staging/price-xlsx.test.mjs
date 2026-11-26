import test from 'node:test';import assert from 'node:assert/strict';import {inspectXlsxZip} from './portal/src/price-xlsx.mjs';
test('rejects invalid and oversized XLSX before ExcelJS parsing',()=>{assert.throws(()=>inspectXlsxZip(new ArrayBuffer(0)));assert.throws(()=>inspectXlsxZip(new ArrayBuffer(5000001)));assert.throws(()=>inspectXlsxZip(new ArrayBuffer(22)));});
