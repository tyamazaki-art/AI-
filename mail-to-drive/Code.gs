/**
 * メール添付ファイルを Google Drive へ自動整理するスクリプト
 * ------------------------------------------------------------
 * 使い方の概要:
 *   1. 専用アドレス（＝あなたのGmail）に、写真やPDFをメールで送る
 *   2. このスクリプトが定期的にメールをチェックする
 *   3. 添付ファイルを種類ごとに Drive の適切なフォルダへ保存する
 *   4. 処理済みのメールにはラベルを付けて二重処理を防ぐ
 *
 * セットアップ手順は README.md を参照してください。
 */

/** ==========================================================
 *  設定エリア（ここだけ自分用に書き換えればOK）
 *  ========================================================== */
const CONFIG = {
  // 添付を保存する「親フォルダ」の名前。Drive内に無ければ自動で作られる。
  ROOT_FOLDER_NAME: '自動整理ボックス',

  // どのルールにも当てはまらなかったファイルの保存先フォルダ名
  DEFAULT_FOLDER: '未分類',

  // 処理が終わったメールに付けるラベル名（Gmailに自動で作られる）
  PROCESSED_LABEL: 'ファイル整理済み',

  // 処理対象のメールを探す条件。
  //   - 特定アドレス宛だけにしたい場合は下の SEARCH_QUERY を
  //     'is:unread has:attachment to:あなた+files@gmail.com' などにする
  SEARCH_QUERY: 'is:unread has:attachment',

  // 1回の実行で処理する最大メール数（多すぎる場合の安全弁）
  MAX_THREADS: 20,

  // 保存が終わったら結果サマリーを自分にメール通知するか
  NOTIFY_SUMMARY: true,

  // 通知の宛先（空にすると実行アカウント本人のアドレスに送る）
  NOTIFY_TO: '',

  /**
   * 振り分けルール。上から順に判定し、最初に当てはまったフォルダへ保存する。
   *   - keywords : ファイル名 or メール件名にこの語が含まれたら該当（大文字小文字は無視）
   *   - extensions: 拡張子がこれに一致したら該当
   * 好きに追加・削除・並べ替えしてください。
   */
  RULES: [
    { folder: '請求書・領収書', keywords: ['請求', '領収', 'invoice', 'receipt'], extensions: [] },
    { folder: '契約・書類',    keywords: ['契約', '見積', '申込', 'contract'],   extensions: [] },
    { folder: '写真',          keywords: [], extensions: ['jpg', 'jpeg', 'png', 'heic', 'gif', 'webp'] },
    { folder: 'PDF書類',       keywords: [], extensions: ['pdf'] },
    { folder: 'オフィス文書',   keywords: [], extensions: ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'] },
    { folder: '音声',          keywords: [], extensions: ['mp3', 'm4a', 'wav', 'aac'] },
    { folder: '動画',          keywords: [], extensions: ['mp4', 'mov', 'avi'] },
  ],
};

/** ==========================================================
 *  メイン処理（トリガーからはこの関数を呼ぶ）
 *  ========================================================== */
function processIncomingMail() {
  const label = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  const rootFolder = getOrCreateFolder_(DriveApp.getRootFolder(), CONFIG.ROOT_FOLDER_NAME);
  const folderCache = {}; // フォルダ名 -> Folder のキャッシュ

  const threads = GmailApp.search(CONFIG.SEARCH_QUERY, 0, CONFIG.MAX_THREADS);
  if (threads.length === 0) {
    Logger.log('処理対象のメールはありませんでした。');
    return;
  }

  const savedItems = []; // 通知用の記録

  threads.forEach(function (thread) {
    const messages = thread.getMessages();
    messages.forEach(function (message) {
      // 既に処理済みラベルが付いたスレッドはスキップ（保険）
      const attachments = message.getAttachments({ includeInlineImages: false });
      if (attachments.length === 0) return;

      const subject = message.getSubject() || '';

      attachments.forEach(function (att) {
        // 実体のない添付（署名画像など極小ファイル）を除外したい場合はここで判定可能
        if (att.getSize() === 0) return;

        const fileName = att.getName() || 'no-name';
        const folderName = decideFolder_(fileName, subject);

        // フォルダを用意（キャッシュ利用）
        if (!folderCache[folderName]) {
          folderCache[folderName] = getOrCreateFolder_(rootFolder, folderName);
        }
        const targetFolder = folderCache[folderName];

        // ファイル名の重複を避けて保存
        const finalName = uniqueFileName_(targetFolder, fileName);
        targetFolder.createFile(att.copyBlob().setName(finalName));

        savedItems.push({ file: finalName, folder: folderName, from: message.getFrom() });
        Logger.log('保存: ' + finalName + ' → ' + CONFIG.ROOT_FOLDER_NAME + '/' + folderName);
      });
    });

    // スレッド全体を処理済みにする
    thread.addLabel(label);
    thread.markRead();
  });

  if (CONFIG.NOTIFY_SUMMARY && savedItems.length > 0) {
    sendSummary_(savedItems);
  }
}

/** ==========================================================
 *  補助関数
 *  ========================================================== */

/** ファイル名・件名から保存先フォルダ名を決める */
function decideFolder_(fileName, subject) {
  const nameLower = String(fileName).toLowerCase();
  const subjectLower = String(subject).toLowerCase();
  const ext = getExtension_(nameLower);

  for (let i = 0; i < CONFIG.RULES.length; i++) {
    const rule = CONFIG.RULES[i];

    // 拡張子一致
    if (rule.extensions && rule.extensions.indexOf(ext) !== -1) {
      return rule.folder;
    }
    // キーワード一致（ファイル名 or 件名）
    if (rule.keywords && rule.keywords.length > 0) {
      for (let k = 0; k < rule.keywords.length; k++) {
        const kw = String(rule.keywords[k]).toLowerCase();
        if (nameLower.indexOf(kw) !== -1 || subjectLower.indexOf(kw) !== -1) {
          return rule.folder;
        }
      }
    }
  }
  return CONFIG.DEFAULT_FOLDER;
}

/** 拡張子（ドット無し・小文字）を返す。無ければ空文字 */
function getExtension_(name) {
  const dot = name.lastIndexOf('.');
  if (dot === -1 || dot === name.length - 1) return '';
  return name.substring(dot + 1);
}

/** 親フォルダ直下に指定名のフォルダを取得。無ければ作成 */
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

/** Gmailラベルを取得。無ければ作成 */
function getOrCreateLabel_(name) {
  const existing = GmailApp.getUserLabelByName(name);
  return existing ? existing : GmailApp.createLabel(name);
}

/** フォルダ内で名前が重複しないよう、必要なら (1)(2)… を付ける */
function uniqueFileName_(folder, name) {
  if (!folder.getFilesByName(name).hasNext()) return name;

  const dot = name.lastIndexOf('.');
  const base = dot === -1 ? name : name.substring(0, dot);
  const ext = dot === -1 ? '' : name.substring(dot); // ドット込み

  let n = 1;
  let candidate;
  do {
    candidate = base + ' (' + n + ')' + ext;
    n++;
  } while (folder.getFilesByName(candidate).hasNext() && n < 1000);
  return candidate;
}

/** 保存結果を自分にメール通知 */
function sendSummary_(items) {
  const to = CONFIG.NOTIFY_TO || Session.getActiveUser().getEmail();
  if (!to) return;

  const lines = items.map(function (it) {
    return '・' + it.file + '  →  ' + CONFIG.ROOT_FOLDER_NAME + '/' + it.folder;
  });

  const body =
    items.length + '件のファイルをDriveへ整理しました。\n\n' +
    lines.join('\n') +
    '\n\n— 自動整理ボックス';

  GmailApp.sendEmail(to, '【自動整理】' + items.length + '件のファイルを保存しました', body);
}

/** ==========================================================
 *  初回セットアップ用: 実行すると15分ごとの自動実行を登録する
 *  （メニューから1回だけ手動実行してください）
 *  ========================================================== */
function setupTrigger() {
  // 既存の同名トリガーを消してから作り直す（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'processIncomingMail') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('processIncomingMail')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('自動実行トリガーを登録しました（15分ごと）。');
}
