/**
 * 네이버 메일 자동 발송 - 진입점
 * 실행: node index.js 또는 npm start
 */

import { getDataFromSheets, sendEmail } from "./tabNaverMail.js";

async function main() {
  try {
    const data = await getDataFromSheets();
    await sendEmail(data);
  } catch (e) {
    console.error("\n=== 오류가 발생했습니다:", e.message, "===");
    process.exitCode = 1;
  }
}

main();
