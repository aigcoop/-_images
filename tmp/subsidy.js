
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
- 사용자가 [수집 자료]로 기사·정부 자료 본문을 줘요. 금액, 대상, 기간, 신청 방법, 시행일, 발표 기관과 날짜는 반드시 이 자료 안에서 확인된 것만 써요. 네가 원래 알던 지식으로 숫자를 채우지 않아요.
- 자료끼리 내용이 다르면 더 최근 자료와 정부 발표를 따르고, checkNeeded에 차이를 적어요.
- 자료에서 확인되지 않는 부분은 숫자를 쓰지 말고 "공고문에서 확인" 식으로 안내하고, checkNeeded에 무엇을 확인해야 하는지 적어요.
- 유튜브 영상 정보는 주제를 찾는 힌트일 뿐이에요. 영상 제목의 과장("수천만 원", "무조건", "속보")을 그대로 쓰지 않아요.
- 수집 자료가 주제와 맞지 않거나, 아직 확정되지 않은 계획·추진 단계 소식뿐이면 skip을 true로 하고 skipReason에 이유를 적어요. 지어내서 쓰는 것보다 보류가 나아요.
- 여러 지자체가 각각 시행하는 지원금이면, 자료에서 확인된 지자체만 표로 정리하고 "지역마다 다르니 주소지 공고 확인"을 강조해요.
- sources와 본문 링크에는 [수집 자료]에 있는 주소만 써요. 공식 원문 버튼은 수집 자료 중 정부·지자체 주소(go.kr, korea.kr 등)를 우선 쓰고, 없으면 관련 기관 누리집 첫 화면(예: https://www.bokjiro.go.kr, https://www.gov.kr, https://www.korea.kr)을 써요.

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
  let hint = input;
  const docs = [];
  if (/youtu/.test(input)) {
    const yt = subYoutube(input);
    hint = (yt && yt.title) + '\n' + (yt && yt.desc);
    ctx += '참고한 유튜브 영상 (주제 힌트로만 사용)\n- 제목: ' + (yt && yt.title) + '\n- 올린 날짜: ' + (yt && yt.date) + '\n- 설명: ' + (yt && yt.desc) + '\n';
  } else if (/^https?:\/\//.test(input)) {
    const t = subFetchText(input, 6000);
    if (t) docs.push({ title: '운영자가 준 공고문·보도자료', url: input, date: '', text: t });
    hint = t ? t.slice(0, 600) : input;
  } else {
    ctx += '주제: ' + input + '\n';
  }
  if (memo) { ctx += '운영자 메모: ' + memo + '\n'; hint += '\n' + memo; }
  const memoUrl = (memo.match(/https?:\/\/\S+/) || [])[0];
  if (memoUrl) { const t = subFetchText(memoUrl, 6000); if (t) docs.push({ title: '운영자가 준 공식 자료', url: memoUrl, date: '', text: t }); }

  const qs = subQueries(hint);
  console.log('검색어: ' + qs.join(' / '));
  subResearch(qs, docs);
  if (!docs.length) return { skip: true, skipReason: '관련 기사나 공식 자료를 찾지 못했어요. 메모 칸에 공고문 주소를 넣고 다시 시도해 주세요.', sources: [] };

  ctx += '\n[수집 자료] (이 안에서 확인된 내용만 쓰세요)\n';
  docs.slice(0, 10).forEach((d, i) => { ctx += '\n### 자료 ' + (i + 1) + ': ' + d.title + '\n주소: ' + d.url + (d.date ? '\n날짜: ' + d.date : '') + '\n' + d.text + '\n'; });
  try {
    const titles = bloggerFetch('/blogs/' + SUB_BLOG + '/posts?maxResults=50&fetchBodies=false&status=live&status=draft&view=ADMIN', 'get').items || [];
    if (titles.length) ctx += '\n이미 있는 글 제목 (같은 주제면 다른 각도로 쓰거나 skip):\n' + titles.map(p => '- ' + p.title).join('\n') + '\n';
  } catch (e) {}

  const txt = callGemini(SUB_SKILL, [{ text: ctx }], null);
  const s = txt.indexOf('{'), e = txt.lastIndexOf('}');
  if (s < 0 || e < 0) throw new Error('제미나이 답에서 JSON을 찾지 못했어요: ' + txt.slice(0, 120));
  const out = JSON.parse(txt.slice(s, e + 1));
  if (out.skip) return out;
  if (!out.title || !out.html) throw new Error('제목이나 본문이 비어 있어요.');
  out.labels = (out.labels || []).filter(l => SUB_LABELS.indexOf(l) !== -1);
  if (!out.labels.length) out.labels = ['생활·안전'];
  out.permalink = String(out.permalink || '').toLowerCase().replace(/[^a-z0-9-]+/g, '-').replace(/^-|-$/g, '');
  const allowed = docs.map(d => d.url);
  out.sources = (out.sources || []).filter(x => x && x.url && (allowed.indexOf(x.url) !== -1 || /\.go\.kr|korea\.kr/.test(x.url)));
  out.html = subFixLinks(out.html);
  out.html = subRelated(out.html, out.labels, out.title);
  return out;
}

/* 영상 제목 등에서 뉴스 검색어 2개 뽑기 */
function subQueries(hint) {
  try {
    const t = callGemini('너는 검색어를 만드는 도우미예요. 설명 없이 JSON 배열만 출력해요.',
      [{ text: '아래 내용이 다루는 한국 정부·지자체 지원 제도를 뉴스에서 찾으려 해요. 과장 표현을 빼고, 제도 이름 중심의 한국어 검색어 2개를 ["검색어1","검색어2"] 형식으로만 답해요.\n\n' + String(hint).slice(0, 1500) }], null);
    const a = JSON.parse(t.slice(t.indexOf('['), t.lastIndexOf(']') + 1));
    if (a.length) return a.slice(0, 2).map(String);
  } catch (e) { console.log('검색어 생성 실패: ' + e.message); }
  return [String(hint).split('\n')[0].replace(/[^\w가-힣0-9 ]/g, ' ').replace(/(속보|정부|꼭|빨리|신청하세요|놀랍습니다)/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 40)];
}

/* 빙 뉴스·빙 검색·구글 뉴스에서 자료 모으기 */
function subResearch(queries, docs) {
  const seen = {};
  docs.forEach(d => seen[d.url] = 1);
  const cutoff = Date.now() - 60 * 24 * 3600 * 1000;
  const items = [];
  queries.forEach(q => {
    subRss('https://www.bing.com/news/search?format=rss&q=' + encodeURIComponent(q)).forEach(x => items.push(x));
    subRss('https://www.bing.com/search?format=rss&q=' + encodeURIComponent(q + ' 정책브리핑 OR 보도자료')).forEach(x => { x.web = true; items.push(x); });
  });
  const fresh = items.filter(x => x.web || !x.date || new Date(x.date).getTime() > cutoff);
  fresh.sort((a, b) => (/\.go\.kr|korea\.kr/.test(b.url) ? 1 : 0) - (/\.go\.kr|korea\.kr/.test(a.url) ? 1 : 0));
  let fetched = 0;
  for (const x of fresh) {
    if (seen[x.url] || docs.length >= 8) continue;
    seen[x.url] = 1;
    let text = '';
    if (fetched < 5) { text = subFetchText(x.url, 3500); fetched++; }
    if (!text || text.length < 300) text = x.desc;
    if (!text) continue;
    docs.push({ title: x.title, url: x.url, date: x.date, text: text });
  }
  // 구글 뉴스 제목 목록 (본문 없이 최근 흐름 참고용)
  try {
    const g = subRss('https://news.google.com/rss/search?hl=ko&gl=KR&ceid=KR:ko&q=' + encodeURIComponent(queries[0] + ' when:30d')).slice(0, 8);
    if (g.length) docs.push({ title: '최근 30일 뉴스 제목 목록 (참고용, 본문 아님)', url: 'https://news.google.com', date: '', text: g.map(x => '- ' + x.title + ' (' + x.date + ')').join('\n') });
  } catch (e) {}
}
function subRss(url) {
  try {
    const t = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'Accept-Language': 'ko-KR' } }).getContentText();
    const dec = s => String(s || '').replace(/<!\[CDATA\[|\]\]>/g, '').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#183;/g, '·').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/<[^>]+>/g, '').trim();
    return t.split('<item>').slice(1, 9).map(x => {
      let link = dec((x.match(/<link>([\s\S]*?)<\/link>/) || [])[1]);
      const u = link.match(/[?&]url=([^&]+)/); if (u) link = decodeURIComponent(u[1]);
      return { title: dec((x.match(/<title>([\s\S]*?)<\/title>/) || [])[1]), url: link, date: dec((x.match(/<pubDate>([\s\S]*?)<\/pubDate>/) || [])[1]), desc: dec((x.match(/<description>([\s\S]*?)<\/description>/) || [])[1]) };
    }).filter(x => x.url);
  } catch (e) { console.log('RSS 실패: ' + url + ' ' + e.message); return []; }
}
function subFetchText(url, max) {
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: false, headers: { 'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ko-KR' } });
    if (r.getResponseCode() >= 400) return '';
    let h = r.getContentText();
    h = h.replace(/<(script|style|noscript|header|footer|nav|aside)[\s\S]*?<\/\1>/gi, ' ');
    const art = h.match(/<article[\s\S]*?<\/article>/i);
    if (art && art[0].length > 1500) h = art[0];
    let t = h.replace(/<br\s*\/?>|<\/p>|<\/li>|<\/tr>|<\/h\d>/gi, '\n').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#\d+;/g, ' ');
    t = t.split('\n').map(s => s.replace(/\s+/g, ' ').trim()).filter(s => s.length > 25 && /[가-힣]/.test(s)).join('\n');
    return t.slice(0, max);
  } catch (e) { return ''; }
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
