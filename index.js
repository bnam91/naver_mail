/**
 * 네이버 메일 자동 발송 - 진입점
 * 실행: node index.js 또는 npm start
 */

import { getDataFromSheets, sendEmail } from "./tabNaverMail.js";

async function main() {
  if (process.stdin.isTTY) {
    process.stdin.resume();
    process.stdin.setRawMode?.(false);
  }
  const isDev = process.argv.includes("--dev");
  if (isDev) console.log("\n🔧 [디버그 모드] 각 동작 후 Enter를 눌러 다음으로 진행합니다.\n");
  try {
    const data = await getDataFromSheets();
    await sendEmail(data, { dev: isDev });
  } catch (e) {
    console.error("\n=== 오류가 발생했습니다:", e.message, "===");
    process.exitCode = 1;
  }
}

main();
