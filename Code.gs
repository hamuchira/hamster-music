/**
 * ハムスターMusicライブ 演奏曲アンケート 受付スクリプト
 *
 * 使い方（README.md にも同じ手順があります）
 * 1. 空のスプレッドシートを作り、「拡張機能 → Apps Script」を開く
 * 2. このファイルの中身をすべて貼り付けて保存
 * 3. 関数 setup を1回だけ実行（権限の確認が出たら許可）
 * 4. 「デプロイ → 新しいデプロイ → ウェブアプリ」
 *      次のユーザーとして実行：自分
 *      アクセスできるユーザー：全員
 * 5. 表示されたウェブアプリのURLを index.html の ENDPOINT に貼る
 */

const MAX_PICKS = 15;
const VOTES = '回答';
const TALLY = '曲別集計';
const RANK = 'ランキング';

// SONGS-START
const SONGS = [
  {cd:"キズナ（関係）", id:"k1", t:"空の下の約束-忘れないで、笑ってて", c:"こふれ", s:122},
  {cd:"キズナ（関係）", id:"k2", t:"空の下の約束-想い続ける、その日まで", c:"ホワイトモカ", s:98},
  {cd:"キズナ（関係）", id:"k3", t:"小娘から母へ", c:"コレット", s:250},
  {cd:"キズナ（関係）", id:"k4", t:"まだそこにいて", c:"ねむりーの", s:419},
  {cd:"キズナ（関係）", id:"k5", t:"巣立ちの朝", c:"とわれ", s:230},
  {cd:"キズナ（関係）", id:"k6", t:"手を伸ばせばそこに -Side S", c:"すふれ", s:277},
  {cd:"キズナ（関係）", id:"k7", t:"手を伸ばせばそこに -Side C", c:"カフェモカ", s:261},
  {cd:"キズナ（関係）", id:"k8", t:"ママは私だけのもの", c:"まどれーぬ", s:261},
  {cd:"キズナ（関係）", id:"k9", t:"真ん中の温度", c:"ソイラテ", s:281},
  {cd:"カオス（世界）", id:"c1", t:"びち太の数え歌 short ver.", c:"びち太", s:28},
  {cd:"カオス（世界）", id:"c2", t:"びち太の数え歌 long ver.", c:"びち太", s:109},
  {cd:"カオス（世界）", id:"c3", t:"僕の方が変態だ", c:"くれーぷ", s:172},
  {cd:"カオス（世界）", id:"c4", t:"フタパカを継ぎし者", c:"しゅぜっと", s:212},
  {cd:"カオス（世界）", id:"c5", t:"パパイヤ", c:"しゅくれ", s:223},
  {cd:"カオス（世界）", id:"c6", t:"奪ってないよ、くれたんだよ", c:"カフェオレ", s:186},
  {cd:"カオス（世界）", id:"c7", t:"土管娘の美容法", c:"さぶれ", s:181},
  {cd:"カオス（世界）", id:"c8", t:"お姉ちゃんにあげる", c:"かぬれ", s:127},
  {cd:"カオス（世界）", id:"c9", t:"私だっていらない", c:"のえる", s:139},
  {cd:"ジブン（個）", id:"j1", t:"アマレッティ・ヤバレッティ", c:"アマレッティ", s:79},
  {cd:"ジブン（個）", id:"j2", t:"アマレッティ・ヤバレッティ -reヤバ-", c:"アマレッティ", s:89},
  {cd:"ジブン（個）", id:"j3", t:"シゴデキ遺伝子を即注入☆", c:"ブラウニー", s:187},
  {cd:"ジブン（個）", id:"j4", t:"Let’sハムササイズ", c:"しふぉん", s:196},
  {cd:"ジブン（個）", id:"j5", t:"たこ焼きわっふる", c:"わっふる", s:133},
  {cd:"ジブン（個）", id:"j6", t:"可愛いは作れる", c:"マキアート", s:120},
  {cd:"ジブン（個）", id:"j7", t:"アイデンティティがクライシス", c:"ぱるふぇ", s:224},
  {cd:"ジブン（個）", id:"j8", t:"一歩だけでえいゆう -real-", c:"らすく", s:215},
  {cd:"ジブン（個）", id:"j9", t:"一歩だけでえいゆう -ideal-", c:"らすく", s:219},
  {cd:"ジブン（個）", id:"j10", t:"宇宙でみんなと交信中", c:"ドルチェ", s:191}
];
// SONGS-END

/** 投票を受け取って「回答」シートに保存する。同じ端末IDからの送り直しは上書き。 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    if (body.hp) return reply_({ ok: true });

    const titleOf = {};
    SONGS.forEach(function (s) { titleOf[s.id] = s.t; });
    const songs = [];
    (body.songs || []).forEach(function (id) {
      if (titleOf[id] && songs.indexOf(id) < 0 && songs.length < MAX_PICKS) songs.push(id);
    });
    if (!songs.length) return reply_({ ok: false, error: 'no songs' });

    const token = String(body.token || '').slice(0, 64);
    const nick = safe_(String(body.nick || '').slice(0, 30));
    const comment = safe_(String(body.comment || '').slice(0, 1000));
    const titles = songs.map(function (id) { return titleOf[id]; }).join('、');
    const row = [new Date(), token, nick, '|' + songs.join('|') + '|', songs.length, titles, comment];

    const sh = SpreadsheetApp.getActive().getSheetByName(VOTES);
    let target = 0;
    if (token && sh.getLastRow() > 1) {
      const hit = sh.getRange(2, 2, sh.getLastRow() - 1, 1)
        .createTextFinder(token).matchEntireCell(true).findNext();
      if (hit) target = hit.getRow();
    }
    if (target) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return reply_({ ok: true });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() {
  return ContentService.createTextOutput('ハムスターMusicライブ アンケート受付中');
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 「=」などで始まる文字がスプレッドシートの数式として動かないようにする。 */
function safe_(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

/** 最初に1回だけ実行：シートと集計の数式を作る。 */
function setup() {
  const ss = SpreadsheetApp.getActive();

  // 回答シート
  let votes = ss.getSheetByName(VOTES) || ss.insertSheet(VOTES);
  votes.clear();
  votes.getRange(1, 1, 1, 7).setValues([['日時', '端末ID', 'ニックネーム', '選曲ID', '曲数', '選んだ曲', 'コメント']]);
  votes.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#E8EAF6');
  votes.setFrozenRows(1);
  votes.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
  votes.setColumnWidth(6, 420);
  votes.setColumnWidth(7, 420);
  votes.getRange('F:G').setWrap(true);
  votes.hideColumns(2, 1);
  votes.hideColumns(4, 1);

  // 曲別集計シート
  let tally = ss.getSheetByName(TALLY) || ss.insertSheet(TALLY);
  tally.clear();
  tally.getRange(1, 1, 1, 6).setValues([['番号', 'ID', '曲名', 'CD', '秒数', '票数']]);
  tally.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#E8EAF6');
  const rows = SONGS.map(function (s, i) { return [i + 1, s.id, s.t, s.cd, s.s]; });
  tally.getRange(2, 1, rows.length, 5).setValues(rows);
  const last = rows.length + 1;
  const counts = [];
  for (let r = 2; r <= last; r++) counts.push(['=COUNTIF(\'' + VOTES + '\'!$D$2:$D,"*|"&B' + r + '&"|*")']);
  tally.getRange(2, 6, rows.length, 1).setFormulas(counts);
  tally.setFrozenRows(1);
  tally.autoResizeColumns(1, 6);

  // ランキングシート（結果はここを見る）
  let rank = ss.getSheetByName(RANK) || ss.insertSheet(RANK);
  rank.clear();
  rank.getRange('A1:B1').setValues([['既存曲に使える時間（分）', 67]]);
  rank.getRange('B1').setBackground('#FFF3C4').setFontWeight('bold');
  rank.getRange('A2').setValue('投票した人');
  rank.getRange('B2').setFormula('=COUNTA(\'' + VOTES + '\'!$A$2:$A)');
  rank.getRange('A3').setValue('コメント数');
  rank.getRange('B3').setFormula('=COUNTIF(\'' + VOTES + '\'!$G$2:$G,"?*")');
  rank.getRange('A5:J5').setValues([['番号', 'ID', '曲名', 'CD', '秒数', '票数', '累計（秒）', '判定', '長さ', '累計']]);
  rank.getRange('A5:J5').setFontWeight('bold').setBackground('#E8EAF6');
  // 票の多い順（同票は元の番号順）。A6 から自動で並ぶ。
  rank.getRange('A6').setFormula('=SORT(\'' + TALLY + '\'!A2:F' + last + ',6,FALSE,1,TRUE)');
  const sub = [];
  for (let r = 6; r < 6 + rows.length; r++) {
    sub.push([
      '=SUM($E$6:E' + r + ')',
      '=IF(G' + r + '<=$B$1*60,"採用","お休み")',
      '=TEXT(E' + r + '/86400,"m:ss")',
      '=TEXT(G' + r + '/86400,"[m]:ss")'
    ]);
  }
  rank.getRange(6, 7, rows.length, 4).setFormulas(sub);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('お休み')
      .setBackground('#FADBD8').setRanges([rank.getRange(6, 8, rows.length, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('採用')
      .setBackground('#D5F0DF').setRanges([rank.getRange(6, 8, rows.length, 1)]).build()
  ];
  rank.setConditionalFormatRules(rules);
  rank.hideColumns(2, 1);
  rank.setFrozenRows(5);
  rank.autoResizeColumns(1, 10);
  rank.setColumnWidth(3, 320);

  // 最初からあるシート1を整理
  const def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  ss.setActiveSheet(rank);
}
