// Render the paperback wrap (PDF) and the Kindle front cover (JPG).
// Usage: node render.js PAGES
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const pages = parseInt(process.argv[2] || '306', 10);
const spine = +(pages * 0.002252).toFixed(3);        // white paper, inches
const bleed = 0.125, trimW = 6, trimH = 9;
const W = bleed * 2 + trimW * 2 + spine, H = trimH + bleed * 2;
const art = fs.readFileSync(path.join(__dirname, 'art.svg'), 'utf8');
const css = fs.readFileSync(path.join(__dirname, 'cover.css'), 'utf8');

function front(x0, y0, s) { // s = inches -> unit
  const u = v => `${(v * s).toFixed(3)}${s === 1 ? 'in' : 'px'}`;
  return `
  <div class="band" style="left:${u(x0)};top:${u(y0 + 7.55)};width:${u(6.0)};height:${u(1.45 + (s===1?bleed:0))}"></div>
  <div class="kicker" style="position:absolute;left:${u(x0+0.55)};top:${u(y0+0.7)};font-size:${u(0.13)}">Eat well · Feel strong · Thrive</div>
  <div class="title" style="position:absolute;left:${u(x0+0.55)};top:${u(y0+1.0)};width:${u(5.0)};font-size:${u(0.62)}">The Menopause Diet for Busy Women</div>
  <div class="subtitle" style="position:absolute;left:${u(x0+0.55)};top:${u(y0+3.05)};width:${u(4.6)};font-size:${u(0.165)}">Easy Recipes, a 12-Week Meal Plan, and Simple Daily Practices to Support Weight Management, Energy, and Wellbeing</div>
  <div style="position:absolute;left:${u(x0+2.35)};top:${u(y0+3.75)};width:${u(3.4)};height:${u(3.4)}">${art.replace('<svg ', '<svg width="100%" height="100%" ')}</div>
  <div style="position:absolute;left:${u(x0+0.55)};top:${u(y0+4.3)};width:${u(2.0)};font-size:${u(0.13)};line-height:1.5;color:#4A4150">
    <div style="font-weight:700;color:#5A3E5D;font-size:${u(0.3)}">30</div>recipes
    <div style="font-weight:700;color:#5A3E5D;font-size:${u(0.3)};margin-top:${u(0.08)}">25</div>snacks
    <div style="font-weight:700;color:#5A3E5D;font-size:${u(0.3)};margin-top:${u(0.08)}">12</div>week meal plan
    <div style="font-weight:700;color:#5A3E5D;font-size:${u(0.3)};margin-top:${u(0.08)}">30</div>days of yoga &amp; strength
  </div>
  <div class="author" style="position:absolute;left:${u(x0)};width:${u(6.0)};text-align:center;top:${u(y0+8.05)};font-size:${u(0.24)}">Neha Menon</div>`;
}

const back = (x0, y0) => `
  <div class="band" style="left:${x0 - bleed}in;top:${y0 + 7.55}in;width:${6.0 + bleed}in;height:${1.45 + bleed}in"></div>
  <div class="blurb" style="position:absolute;left:${x0 + 0.55}in;top:${y0 + 0.6}in;width:4.9in">
    <h2 style="font-size:0.42in">You're not broken.<br>You're changing.</h2>
    <p style="font-size:0.155in;margin-top:0.18in">Waking at 3 a.m. Hot flashes in the middle of a meeting. Cravings, brain fog, and a waistline with a mind of its own. If perimenopause or menopause has made your body feel like a stranger, this book is the friend who did the homework for you.</p>
    <p style="font-size:0.155in;margin-top:0.1in"><i>The Menopause Diet for Busy Women</i> brings together the latest guidelines and research in plain, warm language, then turns them into food and habits that fit a real, busy life.</p>
    <p style="font-size:0.155in;margin-top:0.14in;font-weight:700;color:#5A3E5D">Inside you'll find:</p>
    <ul style="font-size:0.15in;margin-top:0.05in">
      <li>A clear guide to perimenopause, menopause and the years beyond</li>
      <li><b>Anchor, Add, Adapt</b>: a simple three-word method for every meal</li>
      <li>30 easy recipes and 25 simple snacks, each with a protein estimate</li>
      <li>A 12-week meal plan with grocery lists and Sunday prep plans</li>
      <li>A 30-day yoga, strength and calm plan you can do at home</li>
      <li>Honest, evidence-based guidance on hormone therapy and other treatments</li>
    </ul>
    <p style="font-size:0.155in;margin-top:0.14in">No shame, no miracle promises, no cooking for hours. Just the information and support you deserve.</p>
  </div>
  <div style="position:absolute;left:${x0 + 0.55}in;top:${y0 + 7.85}in;width:3.2in;font-size:0.12in;line-height:1.45;color:#F6F1EA">
    <b>Neha Menon</b> is the pen name of a writer committed to giving women clear, kind, evidence-based guidance on menopause.
  </div>
  <div style="position:absolute;left:${x0 + 6.0 - 0.25 - 2.0}in;top:${y0 + 9.0 - 0.25 - 1.2}in;width:2.0in;height:1.2in;background:#FFFFFF"></div>`;

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' }).catch(() => chromium.launch());
  const page = await browser.newPage();
  const xBack = bleed, xSpine = bleed + trimW, xFront = bleed + trimW + spine;
  const wrap = `<html><head><style>${css}</style></head><body>
   <div class="wrap" style="width:${W}in;height:${H}in">
     ${back(xBack, bleed)}
     <div class="band" style="left:${xSpine}in;top:0;width:${spine}in;height:${H}in"></div>
     <div class="spine-text" style="left:${xSpine + spine / 2 - 3.5}in;top:${H / 2 - 0.1}in;width:7in;height:0.2in;text-align:center;font-size:0.17in;font-weight:700">THE MENOPAUSE DIET FOR BUSY WOMEN &nbsp;·&nbsp; <span style="font-weight:500">NEHA MENON</span></div>
     ${front(xFront, bleed, 1)}
     <div class="band" style="left:${xFront + trimW}in;top:0;width:${bleed}in;height:${H}in;background:#F6F1EA"></div>
   </div></body></html>`;
  await page.setContent(wrap, { waitUntil: 'load' });
  await page.pdf({ path: path.join(__dirname, `paperback_cover_${pages}p.pdf`), width: `${W}in`, height: `${H}in`, printBackground: true });
  // Kindle front cover 1600 x 2560 px (1.6:1): scale the 6 x 9 design and extend the band
  const s = 1600 / 6;
  const ebook = `<html><head><style>${css}</style></head><body>
   <div class="wrap" style="width:1600px;height:2560px">
     <div style="position:absolute;left:0;top:0;width:1600px;height:${(0.3*s)}px"></div>
     ${front(0, 0.3, s)}
     <div class="band" style="left:0;top:${(7.85+1.45)*s}px;width:1600px;height:400px"></div>
   </div></body></html>`;
  await page.setViewportSize({ width: 1600, height: 2560 });
  await page.setContent(ebook, { waitUntil: 'load' });
  await page.screenshot({ path: path.join(__dirname, 'kindle_cover.jpg'), type: 'jpeg', quality: 92, clip: { x: 0, y: 0, width: 1600, height: 2560 } });
  await browser.close();
  console.log(`spine ${spine}in, wrap ${W.toFixed(3)} x ${H}in`);
})();
