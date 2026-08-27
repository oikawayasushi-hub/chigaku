/*******************************************************************
 * 地学基礎 確認テスト　記録用スクリプト  v5
 *
 * 【設置手順】
 *  1. Google ドライブで新しいスプレッドシートを作る（名前は何でもよい）
 *  2. メニュー「拡張機能」→「Apps Script」を開く
 *  3. 出てきたコードをすべて消して、このファイルの中身をまるごと貼り付ける
 *  4. 保存（フロッピーのアイコン）
 *  5. 右上の「デプロイ」→「新しいデプロイ」
 *       種類：ウェブアプリ
 *       次のユーザーとして実行：自分
 *       アクセスできるユーザー：全員          ← ここが重要
 *  6. 「デプロイ」を押し、初回は権限の承認を行う
 *       （「このアプリは確認されていません」と出たら
 *         「詳細」→「（プロジェクト名）に移動」→「許可」）
 *  7. 表示された「ウェブアプリのURL」をコピーする
 *       https://script.google.com/macros/s/AKfycb.../exec
 *  8. chigakukiso.html の先頭にある
 *       const SHEET_URL = "";
 *     の "" の中にそのURLを貼り付けて保存する
 *
 * 【コードを直したとき】★ここを忘れると直りません★
 *  「デプロイ」→「デプロイを管理」→ 鉛筆アイコン →
 *  バージョンを「新バージョン」にして「デプロイ」。URLは変わりません。
 *  （「新しいデプロイ」を選ぶとURLが変わってしまうので注意）
 *
 *  保存しただけでは、ウェブアプリは古いコードのまま動き続けます。
 *  いま動いている版は、ブラウザで
 *      （ウェブアプリのURL）?ping=1
 *  を開けば "ver" で確かめられます。v5 と出れば反映されています。
 *
 * 【置き場所について】
 *  学校のGoogle Workspaceでは、管理者の設定によりウェブアプリの公開範囲に
 *  「全員」を選べないことがあります。その場合、個人アカウント（またはこの用途
 *  専用に作った無料アカウント）でスプレッドシートとスクリプトを作れば、
 *  「全員」でデプロイでき、どの端末・どのアカウントからでも記録できます。
 *  記録するのは学籍番号と得点だけで、氏名は含みません。
 *
 * 【v5 での変更点】
 *  ・すべての応答に ver（このスクリプトの版）を入れた。
 *    直したのにデプロイし忘れて古い版が動き続ける、という事故に
 *    その場で気づけるようにするため。
 *  ・verify（記録済みかの確認）を、学籍番号の検証より先に処理するようにした。
 *    確認は読むだけの操作なので、範囲チェックで弾く必要がない。
 *  ・verify が、記録用タブの無いスプレッドシートに対して
 *    タブを新規作成してしまうのをやめた。確認では一切書き込まない。
 *  ・学籍番号が範囲外のときの応答に code:'badsid' を付けた。
 *    通信の失敗と区別できるようにするため。
 *
 * 【v4 での変更点】
 *  ・「記録されたかどうか」をアプリが後から確認できるようにした（verify）。
 *    混雑や通信の乱れで応答が返らなかったときでも、実際に記録されていれば
 *    アプリ側が「送信できた」と正しく表示します。
 *
 * 【v3 での変更点】
 *  ・学籍番号の範囲チェックをサーバ側にも追加（SID_RANGES）。
 *
 * 【v2 での変更点】
 *  ・JSONP（callback パラメータ）に対応。
 *    これにより、HTMLファイルをパソコンに保存してダブルクリックで開いた
 *    状態（file:// ）からでも記録できるようになりました。
 *  ・記録する列を、学籍番号・クラス・問題番号・得点・誤答問題・合否 の形に変更。
 *******************************************************************/

var SCRIPT_VER = 'v5';   // 応答に入れる版。デプロイが反映されたかの確認に使う。
var SHEET_NAME = '記録';
var KEY_COL    = 18;   // 送信キーの列（重複防止に使う）

/* 受け付ける学籍番号の範囲。範囲外からの書き込みは拒否する。
   ウェブアプリのURLはHTMLに書かれていて誰でも見られるため、
   ここでも念のため検証しておく。クラスが増えたらここを直す。 */
var SID_RANGES = [
  [2100, 2699],   // 生徒
  [2000, 2099]    // 教員
];
function sidOK(sid) {
  if (!/^\d{4}$/.test(String(sid))) return false;
  var n = Number(sid);
  for (var i = 0; i < SID_RANGES.length; i++) {
    if (n >= SID_RANGES[i][0] && n <= SID_RANGES[i][1]) return true;
  }
  return false;
}

function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  var cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : '';
  try {
    var p = {};
    if (e && e.postData && e.postData.contents) {
      try { p = JSON.parse(e.postData.contents); }
      catch (err) { p = e.parameter || {}; }
    } else {
      p = (e && e.parameter) ? e.parameter : {};
    }

    // 接続テスト用
    if (p.ping) {
      var ssP = SpreadsheetApp.getActiveSpreadsheet();
      return out({ ok: true, msg: 'connected',
                   sheet: ssP ? ssP.getName() : '(なし)',
                   url:   ssP ? ssP.getUrl()  : '',
                   tab:   SHEET_NAME }, cb);
    }

    // 記録済みかどうかの確認（アプリが送信後の確認に使う）
    // 読むだけで、書き込みは一切しない。記録用タブが無ければ found:false を返す。
    if (p.verify) {
      var shV  = findSheet();                        // 無ければ null（作らない）
      var rowV = shV ? findKeyRow(shV, p.key) : 0;
      return out({ ok: true, found: rowV > 0, row: rowV || '',
                   sheet: SpreadsheetApp.getActiveSpreadsheet().getName(),
                   tab: SHEET_NAME }, cb);
    }

    // 学籍番号の検証（範囲外は記録しない）
    if (!sidOK(p.sid)) {
      return out({ ok: false, code: 'badsid',
                   error: '学籍番号が範囲外です: ' + String(p.sid || '') }, cb);
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var sh = getSheet();

      // 同じ送信キーがすでにあれば追加しない（二重記録の防止）
      var dupRow = findKeyRow(sh, p.key);
      if (dupRow > 0) {
        return out({ ok: true, dup: true, row: dupRow,
                     sheet: SpreadsheetApp.getActiveSpreadsheet().getName(),
                     tab: SHEET_NAME }, cb);
      }

      var n  = Number(p.n  || 0);
      var sc = Number(p.ok || 0);
      sh.appendRow([
        new Date(),                        // 1  送信日時
        String(p.sid || ''),               // 2  学籍番号
        String(p.cls || ''),               // 3  クラス
        String(p.no  || ''),               // 4  出席番号
        p.role  || '',                     // 5  区分（生徒／教員）
        String(p.code || ''),              // 6  問題番号（例 311）
        p.title || '',                     // 7  単元・小単元
        p.mode  || '',                     // 8  種別（確認／復習）
        n,                                 // 9  出題数
        sc,                                // 10 得点
        n ? Math.round(sc / n * 100) : '', // 11 正答率(%)
        p.pass  || '',                     // 12 合否
        p.wrong || '',                     // 13 誤答問題
        p.marks || '',                     // 14 各問の正誤
        p.kcs   || '',                     // 15 項目別の正誤
        Number(p.t || 0),                  // 16 所要秒
        p.ver   || '',                     // 17 アプリ版
        String(p.key || '')                // 18 送信キー
      ]);
      return out({ ok: true, row: sh.getLastRow(),
                   sheet: SpreadsheetApp.getActiveSpreadsheet().getName(),
                   tab: SHEET_NAME }, cb);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out({ ok: false, error: String(err) }, cb);
  }
}

/* 送信キーが記録済みなら、その行番号を返す。なければ 0。 */
function findKeyRow(sh, key) {
  if (!key) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var keys = sh.getRange(2, KEY_COL, last - 1, 1).getValues();
  for (var i = keys.length - 1; i >= 0; i--) {
    if (String(keys[i][0]) === String(key)) return i + 2;
  }
  return 0;
}

/* 記録用のシートタブを返す。無ければ null（作らない）。確認処理から使う。 */
function findSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['送信日時','学籍番号','クラス','出席番号','区分','問題番号','単元・小単元','種別',
                  '出題数','得点','正答率(%)','合否','誤答問題','各問の正誤','項目別の正誤',
                  '所要秒','アプリ版','送信キー']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(7, 240);
    sh.setColumnWidth(13, 200);
    sh.setColumnWidth(15, 300);
    sh.hideColumns(KEY_COL);
  }
  return sh;
}

/* JSONP（callback 付き）なら JavaScript として返す。
   これで file:// から開いたページでも受け取れる。
   どの応答にも ver を入れて、動いている版がすぐ分かるようにしてある。 */
function out(o, cb) {
  o.ver = SCRIPT_VER;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(o) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}


/*******************************************************************
 * 【診断】記録が見つからないときは、この関数を手で実行してください。
 *
 *  1. Apps Script のエディタ上部で、実行する関数を「診断」に切り替える
 *  2. 「実行」を押す
 *  3. 下部の「実行ログ」に、書き込んだスプレッドシートの名前・URL・
 *     シートタブ名・現在の行数が表示されます
 *  4. 表示されたURLを開けば、実際にどのファイルへ記録されているか
 *     一目でわかります（見ていたファイルと違うことがよくあります）
 *
 *  実行すると、テスト用の行が1行追加されます。確認後は消して構いません。
 *******************************************************************/
function 診断() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('★このスクリプトはスプレッドシートに紐づいていません。');
    Logger.log('　スプレッドシートを開き「拡張機能 → Apps Script」から');
    Logger.log('　作り直してください（script.google.com から作った場合はこうなります）。');
    return;
  }
  Logger.log('スクリプトの版     : ' + SCRIPT_VER);
  Logger.log('スプレッドシート名 : ' + ss.getName());
  Logger.log('スプレッドシートURL: ' + ss.getUrl());
  Logger.log('書き込むシートタブ : ' + SHEET_NAME);

  var names = ss.getSheets().map(function (x) { return x.getName(); });
  Logger.log('いまあるシートタブ : ' + names.join(' / '));

  var sh = getSheet();
  sh.appendRow([new Date(), '9999', '9', '99', '診断', '000',
                'テスト書き込み', '診断', 0, 0, '', '', '', '', '', 0, '診断', 'diag-' + Date.now()]);
  Logger.log('テスト行を追加しました。現在の行数 : ' + sh.getLastRow());
  Logger.log('★上のURLを開き、「' + SHEET_NAME + '」タブを見てください。');
}
