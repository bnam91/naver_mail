/**
 * claude_runner.js
 * Claude Code 전용 네이버 메일 발송 헬퍼
 * readline-sync 없이 CLI 인수로 선택값을 받아 실행
 *
 * 사용법:
 *   node claude_runner.js --list
 *   node claude_runner.js --template 1 --db 2 --profile 1
 */

import { createRequire } from "module";
import fs from "fs";
import os from "os";
import path from "path";
import { pathToFileURL } from "url";
import { spawnSync } from "child_process";
import { google } from "googleapis";
import { InstagramMessageTemplate } from "./naverMessageModule.js";

const require = createRequire(import.meta.url);
const { openBrowser } = require("./submodules/module_chrome_set/index.js");

const USER_DATA_PATH = path.join(os.homedir(), "Documents", "github_cloud", "user_data");
const AUTH_PATH = path.join(os.homedir(), "Documents", "github_cloud", "module_auth", "auth.js");
const { getCredentials } = await import(pathToFileURL(AUTH_PATH).href);

const SPREADSHEET_ID = "1yG0Z5xPcGwQs2NRmqZifz0LYTwdkaBwcihheA13ynos";

function log(msg) { console.log(msg); }
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function getNaverProfiles() {
  const profiles = [];
  try {
    for (const item of fs.readdirSync(USER_DATA_PATH)) {
      if (!item.startsWith("naver_")) continue;
      const itemPath = path.join(USER_DATA_PATH, item);
      if (!fs.statSync(itemPath).isDirectory()) continue;
      const hasDefault = fs.existsSync(path.join(itemPath, "Default"));
      const hasProfile = !hasDefault && fs.readdirSync(itemPath).some(s => s.startsWith("Profile"));
      if (hasDefault || hasProfile) profiles.push(item);
    }
  } catch (e) { log(`프로필 읽기 실패: ${e.message}`); }
  return profiles;
}

async function fetchSheetOptions(sheets) {
  const dbRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "아이디보드!F2:F",
  });
  const validDbs = (dbRes.data.values || []).filter(r => r?.[0]).map(r => r[0]);

  const metaRes = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const templateSheets = (metaRes.data.sheets || [])
    .map(s => s.properties?.title || "")
    .filter(t => t && !t.startsWith("_") && t !== "아이디보드" && t !== "사용법" &&
                 !t.includes("완료") && !validDbs.includes(t));

  return { templateSheets, validDbs };
}

// --list 모드: 선택 가능한 목록 출력 후 종료
async function listOptions() {
  const creds = await getCredentials();
  const sheets = google.sheets({ version: "v4", auth: creds });
  const { templateSheets, validDbs } = await fetchSheetOptions(sheets);
  const profiles = getNaverProfiles();

  log("\n=== 템플릿 시트 ===");
  templateSheets.forEach((s, i) => log(`  ${i + 1}. ${s}`));
  log("\n=== DB(메일주소) ===");
  validDbs.forEach((d, i) => log(`  ${i + 1}. ${d}`));
  log("\n=== 네이버 프로필 ===");
  profiles.forEach((p, i) => log(`  ${i + 1}. ${p}`));
  log("");
}

function pasteByPythonOsLevel(text) {
  const bodyFile = path.join(os.tmpdir(), `naver_mail_body_${Date.now()}.txt`);
  fs.writeFileSync(bodyFile, text, "utf8");
  const py = `
import os, sys, time, pyperclip, pyautogui
body_file = os.environ.get("BODY_FILE", "")
if not body_file or not os.path.isfile(body_file): sys.exit(1)
with open(body_file, encoding="utf-8") as f: body = f.read()
try: os.remove(body_file)
except: pass
pyperclip.copy(body)
time.sleep(0.05)
mod = "command" if sys.platform == "darwin" else "ctrl"
pyautogui.keyDown(mod); pyautogui.press("v"); pyautogui.keyUp(mod)
`;
  const pythonCmd = process.platform === "win32" ? "python" : "python3";
  const res = spawnSync(pythonCmd, ["-c", py], {
    env: { ...process.env, BODY_FILE: bodyFile },
    encoding: "utf8",
  });
  try { if (fs.existsSync(bodyFile)) fs.unlinkSync(bodyFile); } catch (_) {}
  if (res.status !== 0) throw new Error((res.stderr || res.stdout || `python exit ${res.status}`).trim());
}

async function getBodyPlainText(page) {
  return page.evaluate(() => {
    const ifr = document.querySelector(".editor_body iframe");
    if (ifr?.contentDocument) {
      const editable = ifr.contentDocument.querySelector("[contenteditable='true']") || ifr.contentDocument.body;
      return (editable?.innerHTML || "").replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "").trim();
    }
    return document.querySelector(".editor_body textarea")?.value?.trim() || "";
  });
}

async function updateSheetStatus(sheets, sheetName, rowIndex, sentTime, status) {
  if (status === "읽지않음") status = "발송완료";
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!D${rowIndex}:E${rowIndex}`,
    valueInputOption: "RAW",
    requestBody: { values: [[sentTime, status]] },
  });
}

// --template N --db N --profile N 모드: 선택값을 받아 발송 실행
async function sendWithSelections(templateIdx, dbIdx, profileIdx) {
  const creds = await getCredentials();
  const sheets = google.sheets({ version: "v4", auth: creds });
  const { templateSheets, validDbs } = await fetchSheetOptions(sheets);
  const profiles = getNaverProfiles();

  const selectedSheet = templateSheets[templateIdx - 1];
  const selectedDb = validDbs[dbIdx - 1];
  const profileName = profiles[profileIdx - 1];

  if (!selectedSheet) { log(`템플릿 번호 ${templateIdx} 없음`); process.exit(1); }
  if (!selectedDb) { log(`DB 번호 ${dbIdx} 없음`); process.exit(1); }
  if (!profileName) { log(`프로필 번호 ${profileIdx} 없음`); process.exit(1); }

  log(`\n선택 완료:`);
  log(`  템플릿: ${selectedSheet}`);
  log(`  DB: ${selectedDb}`);
  log(`  프로필: ${profileName}\n`);

  const titleRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${selectedSheet}!B1:D1`,
  });
  const emailTitles = (titleRes.data.values || [[]])[0];

  const messageTemplate = new InstagramMessageTemplate(SPREADSHEET_ID, selectedSheet);
  const emailContents = await messageTemplate.getMessageTemplates();

  const recipientRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${selectedDb}!B2:D`,
  });
  const allData = recipientRes.data.values || [];
  const recipientData = [];
  let skippedCount = 0;

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const dEmpty = !row[2] || String(row[2]).trim() === "";
    if (dEmpty && row.length >= 2) {
      recipientData.push([row[0], row[1], i + 2]);
      log(`발송 대상: ${row[1]}`);
    } else {
      skippedCount++;
      log(`건너뜀 (이미 발송됨): ${row[1]}`);
    }
  }

  if (skippedCount > 0) log(`\n이미 발송된 ${skippedCount}개 제외`);
  log(`총 ${recipientData.length}개 발송 예정\n`);

  const browser = await openBrowser({
    profileName,
    profilePath: USER_DATA_PATH,
    url: "https://mail.naver.com/v2/new",
    returnBrowser: true,
  });

  if (!browser) { log("브라우저를 열 수 없습니다."); return; }

  const pages = await browser.pages();
  const mailPage = pages.find(p => p.url().includes("mail.naver")) || pages[pages.length - 1];

  try {
    await mailPage.waitForSelector("#recipient_input_element", { timeout: 15000 });

    const total = recipientData.length;
    for (let idx = 0; idx < total; idx++) {
      const [name, email, rowIdx] = recipientData[idx];
      log(`\n--- 메일 ${idx + 1}/${total}: ${name} <${email}> ---`);

      try {
        const htmlBt = await mailPage.$("div.editor_mode_select button[value='HTML']");
        if (htmlBt) { await htmlBt.click(); await sleep(1000); }
      } catch {}

      const addrEl = await mailPage.$("#recipient_input_element");
      await addrEl.click();
      await addrEl.evaluate((el, v) => {
        el.value = v;
        el.dispatchEvent(new Event("input", { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
      }, email);
      await mailPage.keyboard.press("Enter");
      await sleep(1000);

      const emailTitle = emailTitles[Math.floor(Math.random() * emailTitles.length)].replace(/{이름}/g, name);
      const titleEl = await mailPage.$("#subject_title");
      await titleEl.click();
      await titleEl.evaluate((el, v) => {
        el.value = v;
        el.dispatchEvent(new Event("input", { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
      }, emailTitle);
      await mailPage.keyboard.press("Tab");
      await sleep(800);

      const emailContent = emailContents[Math.floor(Math.random() * emailContents.length)].replace(/{이름}/g, name);
      await mailPage.waitForSelector(".editor_body iframe, .editor_body textarea", { timeout: 5000 }).catch(() => null);
      await sleep(300);

      try {
        pasteByPythonOsLevel(emailContent);
        await sleep(500);
        const pasted = await getBodyPlainText(mailPage);
        const probe = emailContent.replace(/\s+/g, " ").slice(0, 8);
        if (!pasted.replace(/\s+/g, " ").includes(probe)) throw new Error("본문 반영 확인 실패");
      } catch (e) {
        log(`  [경고] 본문 입력 실패: ${e.message}`);
      }

      await sleep(2000);
      const sendBtn = await mailPage.$(".button_write_task");
      await sendBtn.click();
      await sleep(5000);

      await mailPage.goto("https://mail.naver.com/v2/folders/2");
      await sleep(5000);

      try {
        const mailItems = await mailPage.$$("li.mail_item.reception");
        if (mailItems.length > 0) {
          const latest = mailItems[0];
          const recipientEl = await latest.$(".recipient_link");
          const recipientText = await recipientEl.evaluate(el => el.textContent);
          const recipientRaw = recipientText.split("\n").pop().trim();
          const recipientMatch = recipientRaw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
          const recipient = recipientMatch ? recipientMatch[0] : recipientRaw;
          let status = await (await latest.$(".sent_status")).evaluate(el => el.textContent);
          let sentTime = await (await latest.$(".sent_time")).evaluate(el => el.textContent);
          const now = new Date();
          sentTime = `${String(now.getFullYear()).slice(-2)}년 ${now.getMonth() + 1}월 ${now.getDate()}일 ${sentTime}`;

          log(`받는사람: ${recipientRaw} | 상태: ${status} | 발송시각: ${sentTime}`);

          if (recipient.toLowerCase() !== email.trim().toLowerCase()) {
            log(`이메일 불일치! 예상: ${email}, 실제: ${recipientRaw}`);
            status = "미발송 (이메일 불일치)"; sentTime = "-";
          }
          await updateSheetStatus(sheets, selectedDb, rowIdx, sentTime, status);
        }
      } catch (e) {
        log(`상태 확인 실패: ${e.message}`);
        await updateSheetStatus(sheets, selectedDb, rowIdx, "-", "미발송 (확인실패)");
      }

      log(`진행상황: ${idx + 1}/${total} 완료`);

      if (idx < total - 1) {
        const waitTime = Math.floor(Math.random() * (70 - 3 + 1)) + 3;
        for (let r = waitTime; r > 0; r--) {
          process.stdout.write(`\r${r}초 후 다음 발송...`);
          await sleep(1000);
        }
        process.stdout.write("\r대기 완료!            \n");
        await mailPage.goto("https://mail.naver.com/v2/new");
        await mailPage.waitForSelector("#recipient_input_element", { timeout: 10000 });
        await sleep(2000);
      }
    }
  } catch (e) {
    log(`\n오류: ${e.message}`);
  } finally {
    log("\n=== 메일 발송 완료 ===");
    await browser.close();
  }
}

// CLI 파싱
const args = process.argv.slice(2);
if (args.includes("--list")) {
  await listOptions().catch(e => { console.error(e.message); process.exit(1); });
} else {
  const tIdx = parseInt(args[args.indexOf("--template") + 1]);
  const dIdx = parseInt(args[args.indexOf("--db") + 1]);
  const pIdx = parseInt(args[args.indexOf("--profile") + 1]);
  if ([tIdx, dIdx, pIdx].some(isNaN)) {
    console.error("사용법:\n  node claude_runner.js --list\n  node claude_runner.js --template 1 --db 2 --profile 1");
    process.exit(1);
  }
  await sendWithSelections(tIdx, dIdx, pIdx).catch(e => { console.error(e.message); process.exit(1); });
}
