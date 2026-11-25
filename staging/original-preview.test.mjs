import test from 'node:test';import assert from 'node:assert/strict';
import {extractOriginalPreview} from './generate-original-preview.mjs';
import {buildOriginalPreviewFrame} from './portal/src/original-preview-frame.mjs';
test('preview extracts original panel and functions only',()=>{const p=extractOriginalPreview();assert.match(p.markup,/id="commessa-form"/);assert.match(p.handlers,/async function createCommessa\(event\)/);assert.match(p.handlers,/function renderCommesseManagementList\(/);assert(!p.handlers.includes('firebase.initializeApp'));});
test('preview network and form submission denied',()=>{const p=buildOriginalPreviewFrame(extractOriginalPreview(),'admin');assert.match(p,/connect-src 'none'/);assert.match(p,/form-action 'none'/);assert.match(p,/driveAccessToken=null/);assert.match(p,/const commesseById=new Map/);assert(!p.includes('src="https://'));});
test('non administrator denied preview creation',()=>assert.throws(()=>buildOriginalPreviewFrame(extractOriginalPreview(),'employee')));
