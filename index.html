/**
 * GAS 本番へコピーする用のソース（この1ファイルを正とする）
 * スプレッドシート ID: 1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E
 * HTML: transfer_form / referral_form / continuation_form / nodai-rugby/index の gasWebAppUrl と対になる Web アプリにデプロイ
 * 農大LP: campaignType「ラグビー割」→ シート「ラグビー割」
 * デプロイ: Webアプリ / 自分で実行 / アクセスは全員
 */

// ▼▼▼【設定項目】▼▼▼ 画像を保存するフォルダの名前
const UPLOAD_FOLDER_NAME = 'JOYFIT乗り換え割_アップロード画像';

/**
 * WebアプリからのPOSTリクエストを処理するメイン関数
 */
function doPost(e) {
  try {
    // データが空の場合のハンドリング
    if (!e || !e.postData || !e.postData.contents) {
        throw new Error('ポストデータが見つかりません。');
    }

    const data = JSON.parse(e.postData.contents);
    const campaignType = data.campaignType;
    
    if (!campaignType) {
      throw new Error('campaignTypeが指定されていません。');
    }

    // ★★★★★★ ここでスプレッドシートIDを直接指定します ★★★★★★
    // 共有いただいたスプレッドシートのID
    const SPREADSHEET_ID = '1w7ExndmZn7t2_z55CvxRDMZy4QAcuEyNhIuj-6sUy3E';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

    const timestamp = new Date();
    let sheet;

    // キャンペーン種別に応じて処理を分岐
    // HTML側の value 値と一致させる必要があります
    switch (campaignType) {
      case '紹介・ペア入会':
      case 'お友達紹介キャンペーン': // 名称変更に対応
        sheet = ss.getSheetByName('紹介・ペア入会');
        if (!sheet) throw new Error('シート「紹介・ペア入会」が見つかりません。');
        // ヘッダー行がない場合は作成
        if (sheet.getLastRow() === 0) {
          sheet.appendRow(["申請日時", "キャンペーン種別", "紹介者_名前", "紹介者_電話番号", "被紹介者_名前", "被紹介者_電話番号", "入会予定店舗"]);
        }
        sheet.appendRow([
          timestamp, campaignType,
          data['ご紹介者様のお名前'], data['ご紹介者様のお電話番号'],
          data['ご友人のお名前'], data['ご友人のお電話番号'],
          data['入会予定店舗']
        ]);
        break;

      case '6ヶ月継続':
      case '6ヶ月継続(2026年8月迄)': // ★HTMLのvalueに合わせて追加しました
        sheet = ss.getSheetByName('6ヶ月継続');
        if (!sheet) throw new Error('シート「6ヶ月継続」が見つかりません。');
        if (sheet.getLastRow() === 0) {
          sheet.appendRow(["申請日時", "キャンペーン種別", "名前", "在籍条件同意"]);
        }
        sheet.appendRow([
          timestamp, campaignType,
          data['お名前'], data['在籍条件同意']
        ]);
        break;

      case '乗り換え':
        sheet = ss.getSheetByName('乗り換え');
        if (!sheet) throw new Error('シート「乗り換え」が見つかりません。');
        if (sheet.getLastRow() === 0) {
          sheet.appendRow(["申請日時", "キャンペーン種別", "名前", "移籍元クラブ", "移籍元地名", "画像リンク"]);
        }
        // 画像保存処理呼び出し
        const imageUrls = saveImagesToDrive(data);
        sheet.appendRow([
          timestamp, campaignType,
          data['お名前'], data['移籍元クラブ'],
          data['移籍元クラブの地名'], imageUrls.join('\n')
        ]);
        break;

      case 'ラグビー割':
        sheet = ss.getSheetByName('ラグビー割');
        if (!sheet) throw new Error('シート「ラグビー割」が見つかりません。');
        if (sheet.getLastRow() === 0) {
          sheet.appendRow(['申請日時', 'キャンペーン種別', 'お名前', 'フリガナ', '学年', '卒業予定年月']);
        }
        sheet.appendRow([
          timestamp,
          campaignType,
          data['お名前'],
          data['フリガナ'],
          data['学年'],
          data['卒業予定年月']
        ]);
        break;

      default:
        // 一致しない場合はエラーログを残しつつ、エラーを返す
        throw new Error('不明なキャンペーン種別です: ' + campaignType);
    }

    // ContentService には withHeaders が無いため setMimeType まで（HTML は no-cors POST で問題になりにくい）
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'データを受信しました。' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log(`[エラー発生] ${error.toString()} | [受信データ] ${e ? e.postData.contents : 'なし'}`);
    
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: `サーバー側でエラーが発生しました: ${error.message}` }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// OPTIONS（必要なら空応答。カスタムヘッダは ContentService では付与不可）
function doOptions(e) {
  return ContentService.createTextOutput('');
}

/**
 * フォルダを取得または作成するヘルパー関数
 */
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

/**
 * 送信されたBase64画像をデコードしてGoogleドライブに保存する関数
 */
function saveImagesToDrive(data) {
  const folder = getOrCreateFolder(UPLOAD_FOLDER_NAME);
  const imageUrls = [];
  const applicantName = data['お名前'] || '不明な申請者';
  
  for (let i = 1; i <= 5; i++) {
    const base64Data = data[`imageBase64_${i}`];
    const fileName = data[`imageFileName_${i}`];
    const mimeType = data[`imageMimeType_${i}`] || 'application/octet-stream';
    
    if (base64Data && fileName) {
      try {
        const decoded = Utilities.base64Decode(base64Data, Utilities.Charset.UTF_8);
        const blob = Utilities.newBlob(decoded, mimeType, fileName);
        const dateStr = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmmss");
        // ファイル名重複防止のためタイムスタンプを付与
        const uniqueFileName = `${dateStr}_${applicantName}_${fileName}`;
        const file = folder.createFile(blob).setName(uniqueFileName);
        
        // 保存したファイルのURLをリストに追加
        imageUrls.push(file.getUrl());
      } catch (e) {
        Logger.log(`[ファイル保存失敗] 申請者: ${applicantName}, ファイル名: ${fileName} | エラー: ${e.toString()}`);
      }
    }
  }
  return imageUrls;
}

/**
 * 権限確認用のダミー関数（初回承認用）
 */
function checkPermissions() {
  try {
    SpreadsheetApp.getActiveSpreadsheet();
    DriveApp.getRootFolder();
    Logger.log('承認プロセスを開始します。');
  } catch (e) {}
}
