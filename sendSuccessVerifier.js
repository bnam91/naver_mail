/**
 * 발송 직후 성공 화면 검수
 * - 성공 상태 래퍼(.mail_status.mail_write_done)
 * - 성공 타이틀 문구("메일을 성공적으로 보냈습니다.")
 * - 필요 시 예상 수신자 이메일 포함 여부 확인
 */

/**
 * @param {import("puppeteer").Page} page
 * @param {{ expectedEmail?: string, timeoutMs?: number }} [options]
 * @returns {Promise<{
 *   ok: boolean,
 *   title?: string,
 *   foundEmails?: string[],
 *   expectedEmail?: string,
 *   message: string
 * }>}
 */
export async function verifySendSuccess(page, options = {}) {
  const { expectedEmail, timeoutMs = 7000 } = options;

  try {
    await page.waitForSelector(".mail_status.mail_write_done", { timeout: timeoutMs });

    const result = await page.evaluate(() => {
      const root = document.querySelector(".mail_status.mail_write_done");
      const title = root?.querySelector(".mail_status_title")?.textContent?.trim() || "";
      const emails = Array.from(root?.querySelectorAll(".user_email") || [])
        .map((el) => el.textContent?.trim() || "")
        .filter(Boolean);

      return { title, emails };
    });

    const titleOk = result.title.includes("메일을 성공적으로 보냈습니다");
    if (!titleOk) {
      return {
        ok: false,
        title: result.title || "(없음)",
        foundEmails: result.emails,
        expectedEmail,
        message: `✗ 발송 성공 화면 문구가 다릅니다: ${result.title || "(없음)"}`,
      };
    }

    if (expectedEmail) {
      const emailOk = result.emails.some((e) => e.toLowerCase() === expectedEmail.toLowerCase());
      if (!emailOk) {
        return {
          ok: false,
          title: result.title,
          foundEmails: result.emails,
          expectedEmail,
          message: `✗ 발송 성공 화면의 이메일 불일치 | 예상: ${expectedEmail} | 실제: ${result.emails.join(", ") || "(없음)"}`,
        };
      }
    }

    return {
      ok: true,
      title: result.title,
      foundEmails: result.emails,
      expectedEmail,
      message: `✓ 발송 성공 화면 검수 통과: ${result.title}`,
    };
  } catch (e) {
    return {
      ok: false,
      expectedEmail,
      message: `✗ 발송 성공 화면 검수 실패: ${e.message}`,
    };
  }
}

