
/* ═════════════ 지원금 노트 반자동화 ═════════════ */
const SUB_SHEET = '지원금';
const SUB_BLOG = '3782791169561383344';
const SUB_HEAD = ['링크 또는 주제', '메모(선택)', '상태', '제목', '퍼머링크', '검색 설명', '라벨', '초안 열기', '생성일', '확인 필요', '참고 출처'];
const SUB_LABELS = ['청년', '중장년·노후', '출산·육아', '소상공인·창업', '주거·생활비', '건강·의료', '생활·안전', '세금·연금'];

function subSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SUB_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SUB_SHEET);
    sh.getRange(1, 1, 1, SUB_HEAD.length).setValues([SUB_HEAD]).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 320); sh.setColumnWidth(2, 180); sh.setColumnWidth(4, 320); sh.setColumnWidth(6, 320); sh.setColumnWidth(10, 260); sh.setColumnWidth(11, 320);
    sh.getRange('A2').setNote('유튜브 영상 링크, 공고문 주소, 또는 "기초연금 신청 방법" 같은 주제를 한 줄에 하나씩 적어요. 상태가 비어 있는 줄부터 위에서 차례로 초안을 만들어요.');
  }
  const st = ss.getSheetByName(SHEET_SETTING);
  const have = st.getDataRange().getValues().map(r => String(r[0]).trim());
  if (have.indexOf('지원금 하루 글 수') === -1) st.appendRow(['지원금 하루 글 수', 2, '지원금 노트 초안을 하루 몇 편 만들지 정해요. 처음 2~3주는 1~2편을 권해요.']);
  if (have.indexOf('지원금 실행 시각') === -1) st.appendRow(['지원금 실행 시각', 7, '이 시각부터 1시간 간격으로 한 편씩 만들어요 (0~23).']);
  return sh;
}
function subHead(sh) { const h = {}; sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach((v, i) => h[String(v).trim()] = i); return h; }
function subSet(sh, row, h, name, val) { if (h[name] !== undefined) sh.getRange(row, h[name] + 1).setValue(val); }

/* 메뉴용 */
function subMakeSheet() { subSheet(); SpreadsheetApp.getUi().alert('지원금 탭을 만들었어요. A열에 영상 링크·공고문 주소·주제를 한 줄에 하나씩 적어 주세요.'); }
function subDraftNow() { const r = subRunOne(true); SpreadsheetApp.getUi().alert(r); }
function subDaily() { subRunOne(false); }
function subEnable() {
  subDisable(true); subSheet();
  const hour = parseInt(getSetting('지원금 실행 시각') || '7', 10);
  for (let i = 0; i < 3; i++) ScriptApp.newTrigger('subDaily').timeBased().everyDays(1).atHour((hour + i) % 24).inTimezone(TZ).create();
  SpreadsheetApp.getUi().alert('매일 ' + hour + '시부터 1시간 간격으로 지원금 초안을 1편씩 만들어요. 하루 ' + (getSetting('지원금 하루 글 수') || '2') + '편이 채워지거나 적어 둔 링크가 없으면 멈춰요.');
}
function subDisable(silent) {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'subDaily') ScriptApp.deleteTrigger(t); });
  if (silent !== true) SpreadsheetApp.getUi().alert('지원금 매일 자동 실행을 껐어요.');
}

/* 한 편 만들기 */
function subRunOne(ignoreQuota) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return '다른 작업이 실행 중이에요. 잠시 후 다시 해 주세요.';
  RUN_STARTED = Date.now();
  try {
    const sh = subSheet();
    const h = subHead(sh);
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const data = sh.getDataRange().getValues();
    const n = parseInt(getSetting('지원금 하루 글 수') || '2', 10);
    const doneToday = data.slice(1).filter(r => String(r[h['상태']]) === '초안 완료' && String(r[h['생성일']]).indexOf(today) === 0).length;
    if (!ignoreQuota && doneToday >= n) return '오늘 만들 초안을 다 만들었어요.';
    let rowNum = 0;
    for (let i = 1; i < data.length; i++) {
      const st = String(data[i][h['상태']]).trim();
      if (String(data[i][h['링크 또는 주제']]).trim() && (st === '' || st === '오류' || st === '작성 중')) { rowNum = i + 1; break; }
    }
    if (!rowNum) return '지원금 탭에 새로 적힌 링크나 주제가 없어요.';
    const row = data[rowNum - 1];
    subSet(sh, rowNum, h, '상태', '작성 중'); SpreadsheetApp.flush();
    try {
      const out = subWrite(String(row[h['링크 또는 주제']]).trim(), String(row[h['메모(선택)']] || '').trim());
      if (out.skip) {
        subSet(sh, rowNum, h, '상태', '보류');
        subSet(sh, rowNum, h, '확인 필요', out.skipReason || '공식 자료로 확인되지 않아 보류했어요.');
        subSet(sh, rowNum, h, '참고 출처', (out.sources || []).map(s => s.url).join('\n'));
        return '보류: ' + (out.skipReason || '');
      }
      const post = bloggerFetch('/blogs/' + SUB_BLOG + '/posts/?isDraft=true', 'post', { kind: 'blogger#post', title: out.title, content: out.html, labels: out.labels });
      const link = 'https://www.blogger.com/blog/post/edit/' + SUB_BLOG + '/' + post.id;
      subSet(sh, rowNum, h, '상태', '초안 완료');
      subSet(sh, rowNum, h, '제목', out.title);
      subSet(sh, rowNum, h, '퍼머링크', out.permalink);
      subSet(sh, rowNum, h, '검색 설명', out.searchDescription);
      subSet(sh, rowNum, h, '라벨', out.labels.join(', '));
      subSet(sh, rowNum, h, '초안 열기', link);
      subSet(sh, rowNum, h, '생성일', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
      subSet(sh, rowNum, h, '확인 필요', out.checkNeeded || '');
      subSet(sh, rowNum, h, '참고 출처', (out.sources || []).map(s => (s.title ? s.title + ' ' : '') + s.url).join('\n'));
      SpreadsheetApp.flush();
      try { subNotice(out, link); } catch (e2) { console.log('알림 실패: ' + e2.message); }
      return '초안 완료: ' + out.title;
    } catch (e) {
      subSet(sh, rowNum, h, '상태', '오류');
      subSet(sh, rowNum, h, '확인 필요', String(e.message).slice(0, 300) + ' (다음 실행 때 다시 시도해요)');
      return '오류: ' + e.message;
    }
  } finally { lock.releaseLock(); }
}

/* 유튜브 영상 제목·설명 읽기 */
function subYoutube(url) {
  const m = url.match(/(?:youtu\.be\/|v=|shorts\/|live\/)([A-Za-z0-9_-]{11})/);
  if (!m) return null;
  const id = m[1];
  let title = '', desc = '', date = '';
  try {
    const html = UrlFetchApp.fetch('https://www.youtube.com/watch?v=' + id + '&hl=ko', { muteHttpExceptions: true, headers: { 'Accept-Language': 'ko-KR,ko;q=0.9' } }).getContentText();
    const vd = html.match(/"videoDetails":\{"videoId":"[^"]+","title":"((?:[^"\\]|\\.)*)"[\s\S]*?"shortDescription":"((?:[^"\\]|\\.)*)"/);
    if (vd) { title = JSON.parse('"' + vd[1] + '"'); desc = JSON.parse('"' + vd[2] + '"'); }
    const pd = html.match(/"publishDate":"([^"]+)"/); if (pd) date = pd[1].slice(0, 10);
  } catch (e) { console.log('유튜브 읽기 실패: ' + e.message); }
  if (!title) {
    try { title = JSON.parse(UrlFetchApp.fetch('https://www.youtube.com/oembed?format=json&url=' + encodeURIComponent('https://www.youtube.com/watch?v=' + id)).getContentText()).title; } catch (e) {}
  }
  desc = desc.split(/정보가 곧 돈|🔴|✅/)[0].trim();
  return { id: id, title: title, desc: desc.slice(0, 1500), date: date };
}

const SUB_SKILL = `너는 구글 블로그 "지원금 노트 - 놓치기 쉬운 정부 혜택 정리"(https://jiwongeum-notes.blogspot.com)의 글을 쓰는 작가예요. 이 블로그는 정부·지자체 지원금과 생활 혜택을 공고문 원문 기준으로 정리하는 개인 블로그이고, 정부 기관이 아니에요.

[가장 중요한 규칙: 사실 확인]
- 반드시 구글 검색으로 정부 공식 자료(정책브리핑 korea.kr, 각 부처·지자체 누리집, 정부24, 복지로, 보조금24, 금융감독원 등)와 주요 언론 보도를 찾아 확인한 내용만 써요.
- 금액, 대상, 기간, 신청 방법, 시행일은 공식 자료에서 확인된 것만 숫자로 써요. 확인이 안 되면 숫자를 쓰지 말고 "공고문에서 확인" 식으로 안내하고, checkNeeded에 무엇을 확인해야 하는지 적어요.
- 유튜브 영상 정보는 주제를 찾는 힌트일 뿐이에요. 영상 제목의 과장("수천만 원", "무조건", "속보")을 그대로 쓰지 않아요.
- 아직 확정되지 않은 계획·추진 단계 소식이거나, 공식 자료를 전혀 찾지 못했다면 skip을 true로 하고 skipReason에 이유를 적어요. 지어내서 쓰는 것보다 보류가 나아요.
- 출처 URL은 검색에서 실제로 확인한 주소만 sources에 넣어요. 주소를 추측해서 만들지 않아요.

[글쓰기]
- 해요체로 처음부터 끝까지 통일, 친절하지만 정확한 설명 톤. 50~70대 독자도 이해하기 쉽게 써요.
- 분량: 공백 제외 2,500~3,500자.
- 핵심키워드는 제목 앞쪽에 넣고, 첫 문단 2문장 안에 1회, 본문 전체에 5~7회 자연스럽게 써요. 그 밖의 같은 단어·어미가 반복되지 않게 표현을 바꿔요.
- 한 문단은 2~4문장으로 짧게.
- 표를 1~3개 넣어요 (대상·금액·기간 정리, 바뀐 점 비교, 신청 단계 등).
- 가르치거나 겁주는 말투, 과장 표현, "무조건 받으세요" 같은 단정은 쓰지 않아요.

[HTML 구조 (순서 고정)]
1. 한눈에 보기 박스: <div style="border:1px solid #d8dee4;border-radius:8px;padding:14px 16px;background:#f6f8fa;margin-bottom:18px;"><strong>한눈에 보기</strong><br>핵심 2~3문장</div>
2. 도입 2문단 (독자 고민 → 이 글에서 무엇을 정리했는지, 발표 기관과 날짜 포함)
3. <!-- 이미지1(대표): 한글로 장면 설명 -->
4. <h2>목차</h2> + <ol>
5. 본문 <h2> 4~6개 (필요하면 <h3>). 중간 한 곳에 <!-- 이미지2(선택): 한글로 장면 설명 -->
6. 핵심 요약 박스 (한눈에 보기와 같은 스타일, <strong>핵심 요약</strong> + <ul> 3~5개)
7. <h2>자주 묻는 질문</h2> + <h3>Q. ...</h3><p>답 2~3문장</p> 3~4개
8. <!-- 관련글 -->
9. <h2>공식 안내 원문 보기</h2> + 버튼 1~2개: <p><a href="공식URL" target="_blank" rel="noopener" style="display:inline-block;padding:10px 16px;background:#1a73e8;color:#fff;border-radius:6px;text-decoration:none;">버튼 문구</a></p> (두 번째 버튼은 background:#34495e)
10. <h2>참고 자료</h2> + <ul> (기관명, 자료명, 날짜)
11. <p style="font-size:13px;color:#666;">이 글은 공개된 공식 안내를 바탕으로 정리한 정보이며, 정부 기관의 공식 입장이 아니에요. ... (작성일 YYYY년 M월 D일)</p>
- 사용 태그: h2, h3, p, ul, ol, li, table, thead, tbody, tr, th, td, strong, a, div, br. h1과 이미지 태그는 쓰지 않아요.
- 표: <table style="width:100%;border-collapse:collapse;">, th는 style="border:1px solid #ccc;padding:8px;background:#f3f4f6;", td는 style="border:1px solid #ccc;padding:8px;".

[출력]
설명 없이 JSON 객체 하나만 출력해요. 키:
title(60자 이내, 핵심키워드 앞쪽), keyword(핵심키워드), permalink(영문 소문자·하이픈, 3~6단어), searchDescription(핵심키워드 포함 120~150자, 해요체), labels(아래 목록에서 1~2개 배열), html(본문 HTML 문자열), sources([{title, url}] 공식 출처 위주 2~5개), checkNeeded(발행 전 사람이 원문과 대조할 항목, 없으면 빈 문자열), skip(true/false), skipReason(문자열).
라벨 목록: ` + SUB_LABELS.join(', ');

function subWrite(input, memo) {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy년 M월 d일');
  let ctx = '오늘 날짜: ' + today + '\n';
  if (/youtu/.test(input)) {
    const yt = subYoutube(input);
    ctx += '참고한 유튜브 영상 (주제 힌트로만 사용)\n- 제목: ' + (yt && yt.title) + '\n- 올린 날짜: ' + (yt && yt.date) + '\n- 설명: ' + (yt && yt.desc) + '\n';
    ctx += '이 영상이 다루는 정부 발표나 공고를 검색으로 찾아, 공식 자료 기준으로 글을 써요.\n';
  } else if (/^https?:\/\//.test(input)) {
    ctx += '공고문·보도자료 주소: ' + input + '\n이 주소의 내용을 읽고, 검색으로 최신 정정 여부와 관련 공식 자료를 함께 확인해 글을 써요.\n';
  } else {
    ctx += '주제: ' + input + '\n이 주제의 최신 공식 발표를 검색으로 찾아 글을 써요.\n';
  }
  if (memo) ctx += '운영자 메모: ' + memo + '\n';
  try {
    const titles = bloggerFetch('/blogs/' + SUB_BLOG + '/posts?maxResults=50&fetchBodies=false&status=live&status=draft&view=ADMIN', 'get').items || [];
    if (titles.length) ctx += '\n이미 있는 글 제목 (같은 주제면 다른 각도로 쓰거나 skip):\n' + titles.map(p => '- ' + p.title).join('\n') + '\n';
  } catch (e) {}

  const txt = subGemini(SUB_SKILL, ctx);
  const s = txt.indexOf('{'), e = txt.lastIndexOf('}');
  if (s < 0 || e < 0) throw new Error('제미나이 답에서 JSON을 찾지 못했어요: ' + txt.slice(0, 120));
  const out = JSON.parse(txt.slice(s, e + 1));
  if (out.skip) return out;
  if (!out.title || !out.html) throw new Error('제목이나 본문이 비어 있어요.');
  out.labels = (out.labels || []).filter(l => SUB_LABELS.indexOf(l) !== -1);
  if (!out.labels.length) out.labels = ['생활·안전'];
  out.permalink = String(out.permalink || '').toLowerCase().replace(/[^a-z0-9-]+/g, '-').replace(/^-|-$/g, '');
  out.html = subFixLinks(out.html);
  out.html = subRelated(out.html, out.labels, out.title);
  return out;
}

/* 검색 기능을 켠 제미나이 호출 */
function subGemini(system, text) {
  const payload = JSON.stringify({
    systemInstruction: { parts: [{ text: system }] },
    contents: [{ role: 'user', parts: [{ text: text }] }],
    tools: [{ google_search: {} }, { url_context: {} }],
    generationConfig: { temperature: 0.5, maxOutputTokens: 32000 }
  });
  const pool = orderedPool();
  let last = '';
  for (let attempt = 0; attempt < 5; attempt++) {
    if (attempt > 0 && Date.now() - RUN_STARTED > 240000) break;
    const model = pool[Math.min(attempt, pool.length - 1)];
    const t0 = Date.now();
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/' + model + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': apiKey() }, payload: payload
    });
    const code = res.getResponseCode();
    console.log('지원금 제미나이 ' + (attempt + 1) + ': ' + model + ' → ' + code + ' (' + Math.round((Date.now() - t0) / 1000) + '초)');
    if (code === 200) {
      const c = (JSON.parse(res.getContentText()).candidates || [])[0];
      const t = c && c.content ? c.content.parts.filter(p => p.text && !p.thought).map(p => p.text).join('') : '';
      if (t.indexOf('{') !== -1) return t;
      last = '빈 답 (' + ((c && c.finishReason) || '') + ')'; continue;
    }
    last = code + ' ' + res.getContentText().slice(0, 150);
    if (code === 429 || code >= 500 || code === 404) { Utilities.sleep(3000); continue; }
    if (code === 400 && /url_context|tool/i.test(last)) { return subGeminiSearchOnly(system, text, model); }
    throw new Error('제미나이 오류 ' + last);
  }
  throw new Error('제미나이가 응답하지 않았어요: ' + last);
}
function subGeminiSearchOnly(system, text, model) {
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': apiKey() },
    payload: JSON.stringify({ systemInstruction: { parts: [{ text: system }] }, contents: [{ role: 'user', parts: [{ text: text }] }], tools: [{ google_search: {} }], generationConfig: { temperature: 0.5, maxOutputTokens: 32000 } })
  });
  if (res.getResponseCode() !== 200) throw new Error('제미나이 오류 ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 150));
  const c = (JSON.parse(res.getContentText()).candidates || [])[0];
  return c.content.parts.filter(p => p.text && !p.thought).map(p => p.text).join('');
}

/* 본문 링크가 실제로 열리는지 확인, 안 열리면 기관 첫 화면으로 */
function subFixLinks(html) {
  return html.replace(/href="(https?:\/\/[^"]+)"/g, (all, url) => {
    try {
      const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: false });
      if (r.getResponseCode() < 400) return all;
    } catch (e) {}
    const root = url.match(/^https?:\/\/[^\/]+/)[0];
    console.log('열리지 않는 링크 → 첫 화면으로 교체: ' + url);
    return 'href="' + root + '"';
  });
}

/* 같은 라벨의 기존 글 2~3개를 하단에 연결 */
function subRelated(html, labels, title) {
  let items = [];
  try {
    labels.forEach(l => {
      const r = bloggerFetch('/blogs/' + SUB_BLOG + '/posts?labels=' + encodeURIComponent(l) + '&maxResults=5&fetchBodies=false', 'get');
      (r.items || []).forEach(p => { if (p.title !== title && !items.some(x => x.url === p.url)) items.push({ title: p.title, url: p.url }); });
    });
    if (items.length < 2) {
      const r = bloggerFetch('/blogs/' + SUB_BLOG + '/posts?maxResults=5&fetchBodies=false', 'get');
      (r.items || []).forEach(p => { if (p.title !== title && !items.some(x => x.url === p.url)) items.push({ title: p.title, url: p.url }); });
    }
  } catch (e) { console.log('관련 글 불러오기 실패: ' + e.message); }
  items = items.slice(0, 3);
  const block = items.length ? '<h2>함께 보면 좋은 글</h2><ul>' + items.map(x => '<li><a href="' + x.url + '">' + x.title + '</a></li>').join('') + '</ul>' : '';
  return html.indexOf('<!-- 관련글 -->') !== -1 ? html.replace('<!-- 관련글 -->', block) : html + block;
}

function subNotice(out, link) {
  const to = getSetting('알림 이메일');
  if (!to) return;
  const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;');
  let html = '<p>지원금 노트에 초안 1편이 준비됐어요.</p><p><b>' + esc(out.title) + '</b> (' + esc(out.labels.join(', ')) + ')<br>'
    + '<a href="' + link + '">초안 바로 열기</a><br>퍼머링크: ' + esc(out.permalink) + '<br>검색 설명: ' + esc(out.searchDescription) + '</p>';
  if (out.checkNeeded) html += '<p style="color:#b45309"><b>발행 전 원문과 꼭 대조할 것:</b><br>' + esc(out.checkNeeded) + '</p>';
  if (out.sources && out.sources.length) html += '<p>참고 출처:<br>' + out.sources.map(s => '<a href="' + s.url + '">' + esc(s.title || s.url) + '</a>').join('<br>') + '</p>';
  html += '<p>대표 이미지 1장을 넣고, 퍼머링크·검색 설명을 붙여넣은 뒤 발행해 주세요.</p>';
  MailApp.sendEmail({ to: to, subject: '[지원금 노트] 초안 준비: ' + out.title, htmlBody: html });
}
