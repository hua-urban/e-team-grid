// E組九宮格 — 空白討論版資料檔（blank.html 專用，與正式 data.js 完全分開）
// 用途：2026-08-04 Joan 現場跟夥伴討論用。8 個空白席位，每席 8 主題 × 8 格全空。
// 安全設計：blank.html 的 API_URL 已清空 → localStorage-only 模式，
//   在這裡打的字只存在當前這台瀏覽器，不會寫回正式版的 Google Sheet 後端。
//   localStorage key 也改成 e-team-blank-* 前綴，與正式版 e-team-grid-* 互不覆蓋。

const BLANK_SEATS = 8;

function makeBlankPerson(n){
  return {
    self:  { name: "空白 " + n, en: "", role: "", area: "" },
    photo: "",
    themes: Array.from({ length: 8 }, function(){
      return { service: "", ta: "", roles: ["","","","","","","",""] };
    })
  };
}

const PEOPLE_DATA = {};
for (let i = 1; i <= BLANK_SEATS; i++) PEOPLE_DATA["b" + i] = makeBlankPerson(i);
