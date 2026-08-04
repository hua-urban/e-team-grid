// E組九宮格 — 空白討論版資料檔（blank.html 專用，與正式 data.js 完全分開）
// 用途：Joan 現場跟夥伴討論用的一頁式空白 81 格。
// 2026-08-04 二版：Joan 要求「只要一頁、不要側邊欄」→ 從 8 席改為單一席位 b1。
// 安全設計：blank.html 的 API_URL 已清空 → localStorage-only 模式，
//   填的字存在當台瀏覽器（關掉重開還在），不會寫回正式版的 Google Sheet 後端。
//   localStorage key 為 e-team-blank-b1-edits-v1，與正式版 e-team-grid-* 互不覆蓋。

const PEOPLE_DATA = {
  "b1": {
    self:  { name: "空白九宮格", en: "", role: "", area: "" },
    photo: "",
    themes: Array.from({ length: 8 }, function(){
      return { service: "", ta: "", roles: ["","","","","","","",""] };
    })
  }
};
