/**
 * 보내기 버튼 클릭 전 입력 검수
 * 받는사람, 제목, 본문이 올바르게 입력됐는지 확인
 */

/**
 * @param {import('puppeteer').Page} page
 * @param {string} expectedEmail
 * @returns {Promise<{ok: boolean, actual?: string, expected: string, message: string}>}
 */
export async function verifyRecipient(page, expectedEmail) {
  try {
    const actual = await page.evaluate(() => {
      const el = document.querySelector("#recipient_input_element");
      return el?.value?.trim() || el?.textContent?.trim() || "";
    });
    const ok = actual.toLowerCase().includes(expectedEmail.toLowerCase());
    return {
      ok,
      actual: actual || "(비어있음)",
      expected: expectedEmail,
      message: ok ? `✓ 받는사람 검수 통과: ${actual}` : `✗ 받는사람 불일치 | 예상: ${expectedEmail} | 실제: ${actual || "(비어있음)"}`,
    };
  } catch (e) {
    return { ok: false, expected: expectedEmail, message: `✗ 받는사람 검수 실패: ${e.message}` };
  }
}

/**
 * @param {import('puppeteer').Page} page
 * @param {string} expectedTitle
 * @returns {Promise<{ok: boolean, actual?: string, expected: string, message: string}>}
 */
export async function verifySubject(page, expectedTitle) {
  try {
    const actual = await page.evaluate(() => {
      const el = document.querySelector("#subject_title");
      return el?.value?.trim() || el?.textContent?.trim() || "";
    });
    const ok = actual === expectedTitle || actual.includes(expectedTitle.slice(0, 20));
    return {
      ok,
      actual: actual || "(비어있음)",
      expected: expectedTitle,
      message: ok ? `✓ 제목 검수 통과: ${actual.slice(0, 50)}${actual.length > 50 ? "..." : ""}` : `✗ 제목 불일치 | 예상: ${expectedTitle.slice(0, 30)}... | 실제: ${actual || "(비어있음)"}`,
    };
  } catch (e) {
    return { ok: false, expected: expectedTitle, message: `✗ 제목 검수 실패: ${e.message}` };
  }
}

/**
 * @param {import('puppeteer').Page} page
 * @param {string} expectedContent - 본문 텍스트 (줄바꿈 포함)
 * @returns {Promise<{ok: boolean, actual?: string, expected: string, message: string}>}
 */
export async function verifyBody(page, expectedContent) {
  try {
    const actual = await page.evaluate(() => {
      const ifr = document.querySelector(".editor_body iframe");
      if (ifr?.contentDocument) {
        const doc = ifr.contentDocument;
        const body = doc.body;
        const editable = doc.querySelector("[contenteditable]") || body;
        const html = editable?.innerHTML || editable?.innerText || editable?.textContent || "";
        return html.replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "").trim();
      }
      const bodyWrap = document.querySelector(".editor_body_wrap, .editor_body, .content_editor_wrap");
      const bodyEd = bodyWrap?.querySelector("[contenteditable='true'], [contenteditable]");
      const all = document.querySelectorAll("[contenteditable='true']");
      const el = bodyEd || (all.length > 1 ? all[all.length - 1] : all[0]);
      const html = el?.innerHTML || el?.innerText || el?.textContent || "";
      return html.replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "").trim();
    });
    const stripHtml = (s) =>
      s.replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "").trim();
    const expectedTrim = stripHtml(expectedContent).trim();
    const actualTrim = actual.trim();
    const sample = expectedTrim.slice(0, 30).replace(/\s+/g, " ");
    const actualSample = actualTrim.slice(0, 30).replace(/\s+/g, " ");
    const ok =
      actualTrim.length > 0 &&
      (actualTrim.includes(sample) ||
        sample.includes(actualSample) ||
        expectedTrim.slice(0, 20) === actualTrim.slice(0, 20));
    return {
      ok,
      actual: actualTrim + (actual.length > 100 ? "..." : ""),
      expected: expectedTrim + "...",
      message: ok ? `✓ 본문 검수 통과 (${actual.length}자)` : `✗ 본문 불일치 | 예상 앞부분: ${expectedTrim.slice(0, 30)}... | 실제 앞부분: ${actualTrim.slice(0, 30) || "(비어있음)"}...`,
    };
  } catch (e) {
    return { ok: false, expected: expectedContent.slice(0, 50), message: `✗ 본문 검수 실패: ${e.message}` };
  }
}

/**
 * 보내기 버튼 클릭 전 전체 입력 검수
 * @param {import('puppeteer').Page} page
 * @param {{email: string, title: string, content: string}} expected
 * @param {(msg: string) => void} log
 * @returns {Promise<{ok: boolean, results: object[]}>} 검수 결과
 */
export async function verifyAllInputs(page, expected, log) {
  log("\n  [검수] 보내기 전 입력 확인 중...");

  const [r, s, b] = await Promise.all([
    verifyRecipient(page, expected.email),
    verifySubject(page, expected.title),
    verifyBody(page, expected.content),
  ]);

  log(`  ${r.message}`);
  log(`  ${s.message}`);
  log(`  ${b.message}`);

  const allOk = r.ok && s.ok && b.ok;
  if (allOk) log("  → 검수 모두 통과");
  else log("  → ⚠ 검수 실패. 보내기를 건너뜁니다.");

  return { ok: allOk, results: [r, s, b] };
}
