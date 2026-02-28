/**
 * naver_message_module.py → Node.js (ESM)
 * 메시지 템플릿 조합 및 변수 치환
 */

import { google } from "googleapis";
import os from "os";
import path from "path";
import { pathToFileURL } from "url";

const AUTH_PATH = path.join(os.homedir(), "Documents", "github_cloud", "module_auth", "auth.js");
const { getCredentials } = await import(pathToFileURL(AUTH_PATH).href);

/**
 * 주어진 날짜의 주차를 계산 (1-5)
 * @param {Date} date
 * @returns {number}
 */
function getWeekNumber(date) {
  const firstDay = new Date(date.getFullYear(), date.getMonth(), 1);
  const firstDayWeekday = firstDay.getDay(); // 0: 일, 1: 월, ..., 6: 토

  const weekNumber = Math.floor((date.getDate() + firstDayWeekday - 1) / 7) + 1;
  return weekNumber;
}

export class InstagramMessageTemplate {
  constructor(templateSheetId, templateSheetName) {
    this.templateSheetId = templateSheetId;
    this.templateSheetName = templateSheetName;
  }

  /**
   * 구글 스프레드시트에서 메시지 템플릿을 가져와 조합
   * @returns {Promise<string[]>}
   */
  async getMessageTemplates() {
    console.log("메시지 템플릿 가져오기 시작");
    try {
      const creds = await getCredentials();
      const sheets = google.sheets({ version: "v4", auth: creds });

      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: this.templateSheetId,
        range: `${this.templateSheetName}!B2:D4`,
      });

      const values = res.data.values || [];

      if (!values || values.length < 3) {
        console.warn("메시지 템플릿을 찾을 수 없습니다.");
        return ["안녕하세요"];
      }

      const pick = (arr) => (arr && arr.length ? arr[Math.floor(Math.random() * arr.length)] : "");
      const greeting = pick(values[0]);
      const proposal = pick(values[1]);
      const closing = pick(values[2]);

      const message = `${greeting}\n\n${proposal}\n\n${closing}`;
      return [message];
    } catch (e) {
      console.error("메시지 템플릿을 가져오는 중 오류 발생:", e);
      return ["안녕하세요"];
    }
  }

  /**
   * 템플릿 변수 치환
   * @param {string} template
   * @param {string} [name=""]
   * @param {string} [notionList=""]
   * @param {string} [totalList=""]
   * @returns {string}
   */
  formatMessage(template, name = "", notionList = "", totalList = "") {
    const now = new Date();
    const weekNumber = getWeekNumber(now);
    const weekDate = `${now.getMonth() + 1}월 ${weekNumber}주차`;

    return template
      .replace(/{이름}/g, name)
      .replace(/{노션리스트}/g, notionList)
      .replace(/{전체리스트}/g, totalList)
      .replace(/{상품리스트}/g, notionList)
      .replace(/{주날짜}/g, weekDate);
  }
}
