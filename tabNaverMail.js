/**
 * tab_naver_mail.py → Node.js (ESM)
 * 네이버 메일 자동 발송
 */

import { Builder, By, Key, until } from "selenium-webdriver";
import chrome from "selenium-webdriver/chrome.js";
import os from "os";
import path from "path";
import { pathToFileURL } from "url";
import clipboard from "clipboardy";
import readlineSync from "readline-sync";
import { google } from "googleapis";
import { InstagramMessageTemplate } from "./naverMessageModule.js";

const AUTH_PATH = path.join(os.homedir(), "Documents", "github_cloud", "module_auth", "auth.js");
const { getCredentials } = await import(pathToFileURL(AUTH_PATH).href);

const SPREADSHEET_ID = "1yG0Z5xPcGwQs2NRmqZifz0LYTwdkaBwcihheA13ynos";

function log(msg) {
  console.log(msg);
}

/**
 * @returns {Promise<{emailTitles: string[], emailContents: string[], userId: string, userPw: string, recipientData: [string, string, number][], selectedDb: string}>}
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

  const accountsRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "아이디보드!A1:D",
  });
  const accounts = accountsRes.data.values || [];

  log("\n=== 사용할 계정 번호를 선택하세요 ===");
  accounts.slice(1).forEach((acc) => {
    const note = acc[3] ? ` (${acc[3]})` : "";
    log(`${acc[0]}. ${acc[1]}${note}`);
  });

  let userId, userPw;
  while (true) {
    const num = readlineSync.question("\n번호를 입력하세요: ");
    const found = accounts.slice(1).find((acc) => String(acc[0]) === String(num));
    if (found) {
      userId = found[1];
      userPw = found[2];
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
    userId,
    userPw,
    recipientData,
    selectedDb,
  };
}

async function createDriver() {
  const options = new chrome.Options();
  options.excludeSwitches("enable-logging", "enable-automation");
  options.addArguments(
    "--log-level=3",
    "--silent",
    "--disable-gpu",
    "--no-sandbox",
    "--disable-dev-shm-usage",
    "--disable-gpu-sandbox",
    "--disable-software-rasterizer",
    "--disable-webgl",
    "--disable-webgl2",
    "--disable-logging",
    "--disable-in-process-stack-traces"
  );

  return new Builder().forBrowser("chrome").setChromeOptions(options).build();
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

async function sendEmail(data) {
  const { emailTitles, emailContents, userId, userPw, recipientData, selectedDb } = data;

  log("\n=== 메일 발송을 시작합니다 ===");
  const driver = await createDriver();
  const creds = await getCredentials();
  const sheets = google.sheets({ version: "v4", auth: creds });

  await driver.manage().window().maximize();
  await driver.get("https://mail.naver.com/v2/new");

  try {
    await driver.wait(until.elementLocated(By.css("#id")), 10000);
    const idEl = await driver.findElement(By.css("#id"));
    await clipboard.write(userId);
    await idEl.sendKeys(Key.chord(Key.CONTROL, "v"));
    await sleep(1000);

    const pwEl = await driver.findElement(By.css("#pw"));
    await clipboard.write(userPw);
    await pwEl.sendKeys(Key.chord(Key.CONTROL, "v"));
    await sleep(1000);

    await driver.findElement(By.css(".btn_login")).click();
    log("\n=== 로그인이 완료되었습니다 ===");

    await driver.wait(until.elementLocated(By.css("#recipient_input_element")), 10000);

    const total = recipientData.length;
    for (let idx = 0; idx < total; idx++) {
      const [name, email, rowIdx] = recipientData[idx];

      try {
        const htmlBt = await driver.wait(
          until.elementLocated(By.css("div.editor_mode_select button[value='HTML']")),
          5000
        );
        await htmlBt.click();
        await sleep(1000);
      } catch {
        // HTML 버튼 없으면 무시
      }

      const addrEl = await driver.findElement(By.css("#recipient_input_element"));
      await addrEl.clear();
      await clipboard.write(email);
      await addrEl.sendKeys(Key.chord(Key.CONTROL, "v"));
      await addrEl.sendKeys(Key.ENTER);
      await sleep(1000);

      const titleEl = await driver.findElement(By.css("#subject_title"));
      await titleEl.clear();
      const emailTitle =
        emailTitles[Math.floor(Math.random() * emailTitles.length)].replace(/{이름}/g, name);
      await clipboard.write(emailTitle);
      await titleEl.sendKeys(Key.chord(Key.CONTROL, "v"));
      await titleEl.sendKeys(Key.TAB);
      await sleep(500);

      const emailContent =
        emailContents[Math.floor(Math.random() * emailContents.length)].replace(/{이름}/g, name);
      await clipboard.write(emailContent);
      const actions = driver.actions({ async: true });
      await actions.keyDown(Key.CONTROL).sendKeys("a").keyUp(Key.CONTROL).perform();
      await actions.keyDown(Key.CONTROL).sendKeys("v").keyUp(Key.CONTROL).perform();
      await sleep(2000);

      const sendBtn = await driver.findElement(By.css(".button_write_task"));
      await sendBtn.click();
      await sleep(5000);

      await driver.get("https://mail.naver.com/v2/folders/2");
      await sleep(5000);

      try {
        const mailItems = await driver.wait(
          until.elementsLocated(By.css("li.mail_item.reception")),
          10000
        );
        if (mailItems.length > 0) {
          const latest = mailItems[0];
          const recipientEl = await latest.findElement(By.css(".recipient_link"));
          const recipientText = await recipientEl.getText();
          const recipient = recipientText.split("\n").pop().trim();
          const statusEl = await latest.findElement(By.css(".sent_status"));
          const timeEl = await latest.findElement(By.css(".sent_time"));
          let status = await statusEl.getText();
          let sentTime = await timeEl.getText();

          const now = new Date();
          const dateStr = `${String(now.getFullYear()).slice(-2)}년 ${now.getMonth() + 1}월 ${now.getDate()}일`;
          sentTime = `${dateStr} ${sentTime}`;

          log("\n=== 메일 발송 상태 ===");
          log(`받는사람: ${recipient}`);
          log(`상태: ${status}`);
          log(`발송시각: ${sentTime}`);

          if (recipient.trim().toLowerCase() !== email.trim().toLowerCase()) {
            log(`이메일 주소 불일치! 예상: ${email}, 실제: ${recipient}`);
            status = "미발송 (이메일 불일치)";
            sentTime = "-";
          }

          await updateSheetStatus(sheets, SPREADSHEET_ID, selectedDb, rowIdx, sentTime, status);
        }
      } catch (e) {
        log(`상태 확인 실패: ${e.message}`);
        await updateSheetStatus(sheets, SPREADSHEET_ID, selectedDb, rowIdx, "-", "미발송 (확인실패)");
      }

      log(`\n=== 진행상황: ${idx + 1}/${total} 완료 ===`);

      if (idx < total - 1) {
        const waitTime = Math.floor(Math.random() * (70 - 3 + 1)) + 3;
        log("\n다음 메일 발송까지 대기...");
        for (let r = waitTime; r > 0; r--) {
          process.stdout.write(`\r${r}초 남음...`);
          await sleep(1000);
        }
        process.stdout.write("\r대기 완료!            \n");

        await driver.get("https://mail.naver.com/v2/new");
        await driver.wait(until.elementLocated(By.css("#recipient_input_element")), 10000);
        await sleep(2000);
      }
    }
  } catch (e) {
    log(`\n=== 오류가 발생했습니다: ${e.message} ===`);
  } finally {
    log("\n=== 메일 발송이 완료되었습니다 ===");
    log("Enter 키를 눌러 브라우저를 종료하세요...");
    readlineSync.question("");
    await driver.quit();
  }
}

export { getDataFromSheets, sendEmail };
