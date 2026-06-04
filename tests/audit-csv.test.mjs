/**
 * Code.gs の純粋関数（csvCell_ / buildAuditCsv_）の単体テスト。
 *
 * Code.gs は Google Apps Script 用ファイルだが、トップレベルでは
 * GAS サービス（SpreadsheetApp 等）を呼ばず、定数と関数宣言のみなので、
 * Node の vm でそのまま評価し、純粋関数だけを取り出してテストできる。
 */
import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const source = fs.readFileSync(path.join(__dirname, "..", "Code.gs"), "utf8");

// GAS グローバルは未定義のままでも、純粋関数を呼ぶ限り評価時にエラーにならない
// （これらは関数本体の中でしか参照されない）。
const sandbox = { console };
vm.createContext(sandbox);
vm.runInContext(source + "\nthis.csvCell_ = csvCell_; this.buildAuditCsv_ = buildAuditCsv_;", sandbox);

const { csvCell_, buildAuditCsv_ } = sandbox;

test("csvCell_ leaves plain values untouched", () => {
  assert.equal(csvCell_("Excel"), "Excel");
});

test("csvCell_ returns empty string for null/undefined", () => {
  assert.equal(csvCell_(null), "");
  assert.equal(csvCell_(undefined), "");
});

test("csvCell_ quotes and escapes commas, quotes and newlines (RFC4180)", () => {
  assert.equal(csvCell_("a,b"), '"a,b"');
  assert.equal(csvCell_('x"y'), '"x""y"');
  assert.equal(csvCell_("l1\nl2"), '"l1\nl2"');
});

test("buildAuditCsv_ starts with a UTF-8 BOM", () => {
  assert.equal(buildAuditCsv_([]).charCodeAt(0), 0xfeff);
});

test("buildAuditCsv_ emits the 8-column header even when empty", () => {
  const header = buildAuditCsv_([]).replace(/^﻿/, "").split("\r\n")[0];
  assert.equal(
    header,
    "timestamp_iso8601,tracking_id,host_app,user_email,shape,top_text,middle_text,bottom_text"
  );
});

test("buildAuditCsv_ tolerates a non-array argument", () => {
  assert.ok(buildAuditCsv_(undefined).includes("timestamp_iso8601"));
  assert.ok(buildAuditCsv_(null).includes("timestamp_iso8601"));
});

test("buildAuditCsv_ renders a row and defaults shape to circle", () => {
  const csv = buildAuditCsv_([
    {
      timestamp: "2026-06-04T00:00:00.000Z",
      trackingId: "a3f1b20e",
      host: "spreadsheet",
      user: "taro.yamada@example.com",
      topText: "承認",
      midText: "2026/06/04",
      bottomText: "山田"
    }
  ]);
  const row = csv.replace(/^﻿/, "").split("\r\n")[1];
  assert.equal(
    row,
    "2026-06-04T00:00:00.000Z,a3f1b20e,spreadsheet,taro.yamada@example.com,circle,承認,2026/06/04,山田"
  );
});

test("buildAuditCsv_ escapes a comma in a field so columns are preserved", () => {
  const csv = buildAuditCsv_([{ topText: "部長, 課長" }]);
  assert.ok(csv.includes('"部長, 課長"'));
});

test("buildAuditCsv_ ends with a trailing CRLF", () => {
  assert.ok(buildAuditCsv_([{ trackingId: "id" }]).endsWith("\r\n"));
});
