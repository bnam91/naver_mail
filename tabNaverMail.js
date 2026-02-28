/**
 * 네이버 메일 자동 발송 (Puppeteer + module_chrome_set)
 * - 시트/DB 선택 → 프로필 선택(서브모듈) → 메일 발송
 * - 로그인 없이 프로필 userData 사용
 */

import { createRequire } from "module";
import fs from "fs";
import os from "os";
import path from "path";
import { pathToFileURL } from "url";
import { spawnSync } from "child_process";
import readlineSync from "readline-sync";
import { google } from "googleapis";
import { InstagramMessageTemplate } from "./naverMessageModule.js";

const require = createRequire(import.meta.url);
const { openBrowser } = require("./submodules/module_chrome_set/index.js");

const USER_DATA_PATH = path.join(os.homedir(), "Documents", "github_cloud", "user_data");

const AUTH_PATH = path.join(os.homedir(), "Documents", "github_cloud", "module_auth", "auth.js");
const { getCredentials } = await import(pathToFileURL(AUTH_PATH).href);

const SPREADSHEET_ID = "1yG0Z5xPcGwQs2NRmqZifz0LYTwdkaBwcihheA13ynos";

function log(msg) {
  console.log(msg);
}

/**
 * @returns {Promise<{emailTitles: string[], emailContents: string[], recipientData: [string, string, number][], selectedDb: string}>}
 */
async function getDataFromSheets() {
  log("\n=== 메일 발송 준비를 시작합니다 ===");
  const creds = await getCredentials();
  const sheets = google.sheets({ version: "v4", auth: creds });

  const dbRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "아이디보드!F2:F",
  });
  const dbList = dbRes.data.values || [];
  const validDbs = dbList.filter((row) => row && row[0]).map((row) => row[0]);

  const metaRes = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetsList = metaRes.data.sheets || [];

  log("\n=== 사용할 메일 템플릿 시트를 선택하세요 ===");
  const templateSheets = [];
  for (const sheet of sheetsList) {
    const title = sheet.properties?.title || "";
    if (
      !title.startsWith("_") &&
      title !== "아이디보드" &&
      title !== "사용법" &&
      !title.includes("완료") &&
      !validDbs.includes(title)
    ) {
      templateSheets.push(title);
      log(`${templateSheets.length}. ${title}`);
    }
  }

  let selectedSheet;
  while (true) {
    const num = parseInt(readlineSync.question("\n번호를 입력하세요: "), 10) - 1;
    if (!isNaN(num) && num >= 0 && num < templateSheets.length) {
      selectedSheet = templateSheets[num];
      break;
    }
    log("올바른 번호를 입력해주세요.");
  }

  log("\n=== 사용할 DB를 선택하세요 ===");
  validDbs.forEach((name, i) => log(`${i + 1}. ${name}`));

  let selectedDb;
  while (true) {
    const num = parseInt(readlineSync.question("\n번호를 입력하세요: "), 10) - 1;
    if (!isNaN(num) && num >= 0 && num < validDbs.length) {
      selectedDb = validDbs[num];
      break;
    }
    log("올바른 번호를 입력해주세요.");
  }

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
    const rowIdx = i + 2;
    const dEmpty = !row[2] || String(row[2]).trim() === "";
    if (dEmpty && row.length >= 2) {
      recipientData.push([row[0], row[1], rowIdx]);
      log(`발송 대상: ${row[1]}`);
    } else {
      skippedCount++;
      log(`이미 발송됨 (건너뜀): ${row[1]}`);
    }
  }

  if (skippedCount > 0) log(`\n이미 발송된 ${skippedCount}개의 메일을 제외했습니다.`);
  log(`총 ${recipientData.length}개의 메일을 발송할 예정입니다.\n`);
  log("\n=== 메일 발송 준비가 완료되었습니다 ===");

  return {
    emailTitles,
    emailContents,
    recipientData,
    selectedDb,
  };
}

async function updateSheetStatus(sheets, spreadsheetId, sheetName, rowIndex, sentTime, status) {
  if (status === "읽지않음") status = "발송완료";
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!D${rowIndex}:E${rowIndex}`,
    valueInputOption: "RAW",
    requestBody: { values: [[sentTime, status]] },
  });
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function pasteByPythonOsLevel(text) {
  // Windows에서는 환경변수로 한글 전달 시 인코딩 깨짐 → 임시 파일로 전달
  const bodyFile = path.join(os.tmpdir(), `naver_mail_body_${Date.now()}.txt`);
  fs.writeFileSync(bodyFile, text, "utf8");

  const py = `
import os
import sys
import time
import pyperclip
import pyautogui

body_file = os.environ.get("BODY_FILE", "")
if not body_file or not os.path.isfile(body_file):
    sys.exit(1)
with open(body_file, encoding="utf-8") as f:
    body = f.read()
try:
    os.remove(body_file)
except Exception:
    pass

pyperclip.copy(body)
time.sleep(0.05)

# hotkey()가 Windows에서 Ctrl 없이 'v'만 보내는 경우 있음 → keyDown/press/keyUp 사용
mod = "command" if sys.platform == "darwin" else "ctrl"
try:
    pyautogui.keyDown(mod)
    pyautogui.press("v")
    pyautogui.keyUp(mod)
except Exception as e:
    sys.exit(2)
`;

  // Windows는 보통 python, 맥/리눅스는 python3
  const pythonCmd = process.platform === "win32" ? "python" : "python3";
  const res = spawnSync(pythonCmd, ["-c", py], {
    env: { ...process.env, BODY_FILE: bodyFile },
    encoding: "utf8",
  });

  try {
    if (fs.existsSync(bodyFile)) fs.unlinkSync(bodyFile);
  } catch (_) {}

  if (res.status !== 0) {
    const msg = (res.stderr || res.stdout || "").trim() || `python exit code ${res.status}`;
    throw new Error(msg);
  }
}

async function getBodyPlainText(page) {
  return page.evaluate(() => {
    const ifr = document.querySelector(".editor_body iframe");
    if (ifr?.contentDocument) {
      const doc = ifr.contentDocument;
      const body = doc.body;
      const editable = doc.querySelector("[contenteditable='true'], [contenteditable]") || body;
      const html = editable?.innerHTML || editable?.innerText || editable?.textContent || "";
      return html.replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "").trim();
    }
    const textarea = document.querySelector(".editor_body textarea");
    return textarea?.value?.trim() || "";
  });
}

/** naver_ 프로필 목록 조회 후 readlineSync로 선택 (서브모듈 readline 충돌 방지) */
function selectNaverProfile() {
  const profiles = [];
  try {
    const items = fs.readdirSync(path.join(USER_DATA_PATH));
    for (const item of items) {
      if (!item.startsWith("naver_")) continue;
      const itemPath = path.join(USER_DATA_PATH, item);
      const stat = fs.statSync(itemPath);
      if (!stat.isDirectory()) continue;
      const hasDefault = fs.existsSync(path.join(itemPath, "Default"));
      const hasProfile = !hasDefault && fs.readdirSync(itemPath).some((s) => s.startsWith("Profile"));
      if (hasDefault || hasProfile) profiles.push(item);
    }
  } catch (e) {
    log(`프로필 목록 읽기 실패: ${e.message}`);
    return null;
  }
  if (profiles.length === 0) {
    log("사용 가능한 naver_ 프로필이 없습니다.");
    return null;
  }
  log("\n=== 네이버 프로필을 선택하세요 ===");
  profiles.forEach((p, i) => log(`${i + 1}. ${p}`));
  while (true) {
    const num = parseInt(readlineSync.question("\n번호를 입력하세요: "), 10) - 1;
    if (!isNaN(num) && num >= 0 && num < profiles.length) return profiles[num];
    log("올바른 번호를 입력해주세요.");
  }
}

function stepWait(dev, msg, needEnter = false) {
  if (!dev) return;
  log(`  → ${msg}`);
  if (needEnter) readlineSync.question("  [Enter]: ");
}

async function sendEmail(data, options = {}) {
  const { emailTitles, emailContents, recipientData, selectedDb } = data;
  const { dev = false } = options;

  const profileName = selectNaverProfile();
  if (!profileName) return;

  const browser = await openBrowser({
    profileName,
    profilePath: USER_DATA_PATH,
    url: "https://mail.naver.com/v2/new",
    returnBrowser: true,
  });

  if (!browser) {
    log("브라우저를 열 수 없습니다.");
    return;
  }

  const pages = await browser.pages();
  const mailPage = pages.find((p) => p.url().includes("mail.naver")) || pages[pages.length - 1];

  const creds = await getCredentials();
  const sheets = google.sheets({ version: "v4", auth: creds });

  try {
    await mailPage.waitForSelector("#recipient_input_element", { timeout: 15000 });
    stepWait(dev, "메일 쓰기 페이지 로드했어요. 받는사람 입력란 보이시나요? 제대로 됐으면 엔터 쳐주세요");

    const total = recipientData.length;
    for (let idx = 0; idx < total; idx++) {
      const [name, email, rowIdx] = recipientData[idx];
      if (dev) log(`\n--- 메일 ${idx + 1}/${total}: ${name} <${email}> ---`);

      try {
        const htmlBt = await mailPage.$("div.editor_mode_select button[value='HTML']");
        if (htmlBt) {
          await htmlBt.click();
          await sleep(1000);
          stepWait(dev, "HTML 탭 클릭했어요. 제대로 됐는지 확인해주세요");
        } else {
          stepWait(dev, "HTML 탭 없어서 건너뜀. 확인했으면 엔터 쳐주세요");
        }
      } catch {
        stepWait(dev, "HTML 탭 클릭 실패해서 건너뜀. 확인했으면 엔터 쳐주세요");
      }

      const addrEl = await mailPage.$("#recipient_input_element");
      await addrEl.click();
      await addrEl.evaluate((el, value) => {
        el.value = value;
        el.dispatchEvent(new Event("input", { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
      }, email);
      await mailPage.keyboard.press("Enter");
      await sleep(1000);
      stepWait(dev, "받는사람 입력했어요. 제대로 됐는지 확인해주세요");

      const emailTitle =
        emailTitles[Math.floor(Math.random() * emailTitles.length)].replace(/{이름}/g, name);
      const titleEl = await mailPage.$("#subject_title");
      await titleEl.click();
      await titleEl.evaluate((el, value) => {
        el.value = value;
        el.dispatchEvent(new Event("input", { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
      }, emailTitle);
      await mailPage.keyboard.press("Tab");
      await sleep(800);
      stepWait(dev, "제목 입력했어요. 제대로 됐는지 확인해주세요");

      const emailContent =
        emailContents[Math.floor(Math.random() * emailContents.length)].replace(/{이름}/g, name);
      const bodyContent = dev ? "테스트입니다1" : emailContent;

      await mailPage.waitForSelector(".editor_body iframe, .editor_body textarea", {
        timeout: 5000,
      }).catch(() => null);
      await sleep(300);

      const methodBodyContent = bodyContent;
      let bodyInserted = false;
      try {
        pasteByPythonOsLevel(methodBodyContent);
        await sleep(500);
        const pasted = await getBodyPlainText(mailPage);
        const probe = methodBodyContent.replace(/\s+/g, " ").slice(0, 8);
        bodyInserted = pasted.replace(/\s+/g, " ").includes(probe);
        if (!bodyInserted) throw new Error("python 붙여넣기 후 본문 반영 확인 실패");
        if (dev) log("  [본문] 방법 1 성공: 1) [레거시] Tab 후 평문 클립보드 + Ctrl+V");
      } catch (e) {
        if (dev) log("  [본문] 방법 1 실패: 1) [레거시] Tab 후 평문 클립보드 + Ctrl+V");
      }

      if (!bodyInserted) log("  [경고] 본문 입력 실패");
      await sleep(2000);
      stepWait(dev, "본문 입력했어요. 제대로 됐는지 확인해주세요");

      // stepWait(dev, "보내기 직전입니다. 확인 후 엔터를 눌러주세요", true);

      const sendBtn = await mailPage.$(".button_write_task");
      await sendBtn.click();
      await sleep(5000);
      stepWait(dev, "보내기 버튼 클릭했어요. 제대로 됐는지 확인해주세요");

      await mailPage.goto("https://mail.naver.com/v2/folders/2");
      await sleep(5000);
      stepWait(dev, "보낸편지함으로 이동했어요. 제대로 됐는지 확인해주세요");

      try {
        const mailItems = await mailPage.$$("li.mail_item.reception");
        if (mailItems.length > 0) {
          const latest = mailItems[0];
          const recipientEl = await latest.$(".recipient_link");
          const recipientText = await recipientEl.evaluate((el) => el.textContent);
          const recipientRaw = recipientText.split("\n").pop().trim();
          const recipientMatch = recipientRaw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
          const recipient = recipientMatch ? recipientMatch[0] : recipientRaw;
          const statusEl = await latest.$(".sent_status");
          const timeEl = await latest.$(".sent_time");
          let status = await statusEl.evaluate((el) => el.textContent);
          let sentTime = await timeEl.evaluate((el) => el.textContent);

          const now = new Date();
          const dateStr = `${String(now.getFullYear()).slice(-2)}년 ${now.getMonth() + 1}월 ${now.getDate()}일`;
          sentTime = `${dateStr} ${sentTime}`;

          log("\n=== 메일 발송 상태 ===");
          log(`받는사람: ${recipientRaw}`);
          log(`상태: ${status}`);
          log(`발송시각: ${sentTime}`);

          if (recipient.toLowerCase() !== email.trim().toLowerCase()) {
            log(`이메일 주소 불일치! 예상: ${email}, 실제: ${recipientRaw}`);
            status = "미발송 (이메일 불일치)";
            sentTime = "-";
          }

          await updateSheetStatus(sheets, SPREADSHEET_ID, selectedDb, rowIdx, sentTime, status);
        }
      } catch (e) {
        log(`상태 확인 실패: ${e.message}`);
        await updateSheetStatus(sheets, SPREADSHEET_ID, selectedDb, rowIdx, "-", "미발송 (확인실패)");
      }
      stepWait(dev, "발송 상태 확인했어요. 제대로 됐는지 확인해주세요");

      log(`\n=== 진행상황: ${idx + 1}/${total} 완료 ===`);
      stepWait(dev, "다음 메일 준비했어요. 제대로 됐는지 확인해주세요");

      if (idx < total - 1) {
        const waitTime = dev ? 2 : Math.floor(Math.random() * (70 - 3 + 1)) + 3;
        log(dev ? "\n[디버그] 2초 대기 후 다음 메일..." : "\n다음 메일 발송까지 대기...");
        for (let r = waitTime; r > 0; r--) {
          process.stdout.write(`\r${r}초 남음...`);
          await sleep(1000);
        }
        process.stdout.write("\r대기 완료!            \n");

        await mailPage.goto("https://mail.naver.com/v2/new");
        await mailPage.waitForSelector("#recipient_input_element", { timeout: 10000 });
        await sleep(2000);
        stepWait(dev, "새 메일 쓰기 페이지로 이동했어요. 제대로 됐는지 확인해주세요");
      }
    }
  } catch (e) {
    log(`\n=== 오류가 발생했습니다: ${e.message} ===`);
  } finally {
    log("\n=== 메일 발송이 완료되었습니다 ===");
    if (dev) {
      log("[dev] 브라우저를 열어둡니다.");
    } else {
      log("Enter 키를 눌러 브라우저를 종료하세요...");
      readlineSync.question("");
      await browser.close();
    }
  }
}

export { getDataFromSheets, sendEmail };
