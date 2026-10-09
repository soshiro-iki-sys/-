// guidebook.html を A4 の PDF と、各ページの確認用 PNG に書き出す
const { chromium } = require("playwright");
const path = require("path");
(async () => {
  const dir = process.argv[2];
  const out = process.argv[3];
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 794, height: 1123 }, deviceScaleFactor: 1 });
  await page.goto("file://" + path.join(dir, "guidebook.html"), { waitUntil: "networkidle" });
  await page.evaluate(() => document.fonts.ready);
  await page.pdf({ path: out, format: "A4", printBackground: true, preferCSSPageSize: true });
  // 各ページの溢れチェック
  const over = await page.evaluate(() => [...document.querySelectorAll(".page")].map((p, i) => {
    const pb = p.getBoundingClientRect().bottom;
    const foot = p.querySelector(".foot");
    const limit = foot ? foot.getBoundingClientRect().top : pb;
    let maxB = 0;
    p.querySelectorAll(".body *").forEach((e) => { const b = e.getBoundingClientRect().bottom; if (b > maxB) maxB = b; });
    return { page: i + 1, gap_px: Math.round(limit - maxB) };
  }));
  console.log(JSON.stringify(over));
  await browser.close();
})();
