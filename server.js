// サーバーサイドの処理（Node.js + Express を想定）

// 必要なモジュールのインポート
const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const { check, validationResult } = require('express-validator');
const fs = require('fs');
const path = require('path');
const csv = require('fast-csv');
const ExcelJS = require('exceljs');
const { v4: uuidv4 } = require('uuid');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');

// アプリケーションの初期化
const app = express();

// ミドルウェアの設定
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(helmet()); // セキュリティ強化

// レート制限（DoS攻撃対策）
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15分
  max: 100, // IPアドレスごとに15分間で100リクエストまで
  message: 'リクエスト数が多すぎます。しばらく経ってから再度お試しください。'
});
app.use('/api/submit', limiter);

// CORS設定
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

// バリデーションルールの定義
const applicationValidationRules = [
  check('email').isEmail().withMessage('有効なメールアドレスを入力してください'),
  check('phoneNumber').matches(/^[0-9\-\(\)（）\s]+$/).withMessage('有効な電話番号を入力してください'),
  // 申込種別によって異なるバリデーション
  check('applicationType').custom((value, { req }) => {
    if (value === 'individual') {
      if (!req.body.fullName) throw new Error('氏名を入力してください');
      if (!req.body.furigana) throw new Error('フリガナを入力してください');
    } else if (value === 'corporate') {
      if (!req.body.companyName) throw new Error('会社名を入力してください');
      if (!req.body.contactPerson) throw new Error('担当者名を入力してください');
    } else {
      throw new Error('申込種別を選択してください');
    }
    return true;
  }),
  check('eventType').not().isEmpty().withMessage('参加イベントを選択してください'),
  check('participationDate').not().isEmpty().withMessage('参加希望日を選択してください'),
  check('numberOfPeople').not().isEmpty().withMessage('参加人数を選択してください'),
  check('agree').equals('true').withMessage('プライバシーポリシーへの同意が必要です')
];

// 申込データを保存するデータベース（実際の実装ではデータベースを使用）
let applications = [];

// 申込データの保存
function saveApplication(application) {
  // 申込IDを生成
  application.id = uuidv4();
  
  // 申込日時を追加
  application.createdAt = new Date();
  
  // 申込ステータスを「保留中」に設定
  application.status = 'pending';
  
  // データベースに保存（この例では配列に追加）
  applications.push(application);
  
  return application;
}

// 申込ステータスの更新
function updateApplicationStatus(id, status) {
  const index = applications.findIndex(app => app.id === id);
  
  if (index !== -1) {
    applications[index].status = status;
    applications[index].updatedAt = new Date();
    return applications[index];
  }
  
  return null;
}

// メール送信設定
const transporter = nodemailer.createTransport({
  host: 'smtp.example.com',
  port: 587,
  secure: false,
  auth: {
    user: 'noreply@example.com',
    pass: 'password'
  }
});

// 申込者への自動返信メール送信
function sendConfirmationEmail(application) {
  // メールアドレスが無効な場合は送信しない
  if (!application.email) {
    console.error('Invalid email address');
    return;
  }
  
  // メールテンプレートの作成
  const mailOptions = {
    from: '"イベント事務局" <noreply@example.com>',
    to: application.email,
    subject: '【イベント申込】お申し込み受付のお知らせ',
    text: `
${application.applicationType === 'individual' ? application.fullName : application.contactPerson} 様

この度はお申し込みいただき、誠にありがとうございます。
以下の内容で申込を受け付けました。

■申込内容
申込ID: ${application.id}
申込種別: ${application.applicationType === 'individual' ? '個人' : '法人'}
${application.applicationType === 'individual' ? 
  `氏名: ${application.fullName}
フリガナ: ${application.furigana}` : 
  `会社名: ${application.companyName}
部署名: ${application.department || 'なし'}
担当者名: ${application.contactPerson}`}
メールアドレス: ${application.email}
電話番号: ${application.phoneNumber}

■イベント情報
参加イベント: ${application.eventType === 'seminar' ? 'セミナー' : application.eventType === 'workshop' ? 'ワークショップ' : 'カンファレンス'}
参加希望日: ${application.participationDate}
参加人数: ${application.numberOfPeople === '5' ? application.exactNumber + '名' : application.numberOfPeople + '名'}

このメールは自動送信されています。
ご不明な点がございましたら、下記までお問い合わせください。

イベント事務局
TEL: 03-1234-5678
Email: event@example.com
    `,
    html: `
<div style="font-family: sans-serif; max-width: 600px; margin: 0 auto;">
  <h2 style="color: #4a6da7;">お申し込み受付のお知らせ</h2>
  <p>${application.applicationType === 'individual' ? application.fullName : application.contactPerson} 様</p>
  <p>この度はお申し込みいただき、誠にありがとうございます。<br>以下の内容で申込を受け付けました。</p>
  
  <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 20px 0;">
    <h3 style="color: #4a6da7; margin-top: 0;">申込内容</h3>
    <p>申込ID: <strong>${application.id}</strong></p>
    <p>申込種別: <strong>${application.applicationType === 'individual' ? '個人' : '法人'}</strong></p>
    ${application.applicationType === 'individual' ? 
      `<p>氏名: <strong>${application.fullName}</strong></p>
      <p>フリガナ: <strong>${application.furigana}</strong></p>` : 
      `<p>会社名: <strong>${application.companyName}</strong></p>
      <p>部署名: <strong>${application.department || 'なし'}</strong></p>
      <p>担当者名: <strong>${application.contactPerson}</strong></p>`}
    <p>メールアドレス: <strong>${application.email}</strong></p>
    <p>電話番号: <strong>${application.phoneNumber}</strong></p>
  </div>
  
  <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 20px 0;">
    <h3 style="color: #4a6da7; margin-top: 0;">イベント情報</h3>
    <p>参加イベント: <strong>${application.eventType === 'seminar' ? 'セミナー' : application.eventType === 'workshop' ? 'ワークショップ' : 'カンファレンス'}</strong></p>
    <p>参加希望日: <strong>${application.participationDate}</strong></p>
    <p>参加人数: <strong>${application.numberOfPeople === '5' ? application.exactNumber + '名' : application.numberOfPeople + '名'}</strong></p>
  </div>
  
  <p>このメールは自動送信されています。<br>ご不明な点がございましたら、下記までお問い合わせください。</p>
  
  <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
    <p style="margin: 0;">イベント事務局</p>
    <p style="margin: 0;">TEL: 03-1234-5678</p>
    <p style="margin: 0;">Email: event@example.com</p>
  </div>
</div>
    `
  };
  
  // メール送信
  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.error('Error sending email:', error);
    } else {
      console.log('Email sent:', info.response);
    }
  });
}

// 事務局への通知メール送信
function sendNotificationEmail(application) {
  const mailOptions = {
    from: '"イベントシステム" <noreply@example.com>',
    to: 'event@example.com',
    subject: '【新規申込】イベント申込がありました',
    text: `
新規申込がありました。

■申込内容
申込ID: ${application.id}
申込日時: ${application.createdAt.toLocaleString('ja-JP')}
申込種別: ${application.applicationType === 'individual' ? '個人' : '法人'}
${application.applicationType === 'individual' ? 
  `氏名: ${application.fullName}
フリガナ: ${application.furigana}` : 
  `会社名: ${application.companyName}
部署名: ${application.department || 'なし'}
担当者名: ${application.contactPerson}`}
メールアドレス: ${application.email}
電話番号: ${application.phoneNumber}

■イベント情報
参加イベント: ${application.eventType === 'seminar' ? 'セミナー' : application.eventType === 'workshop' ? 'ワークショップ' : 'カンファレンス'}
参加希望日: ${application.participationDate}
参加人数: ${application.numberOfPeople === '5' ? application.exactNumber + '名' : application.numberOfPeople + '名'}

■備考
${application.notes || 'なし'}

管理画面から詳細を確認してください。
    `
  };
  
  // メール送信
  transporter.sendMail(mailOptions);
}

// 申込データのCSVエクスポート
function exportToCsv() {
  return new Promise((resolve, reject) => {
    const filename = `applications_${Date.now()}.csv`;
    const csvStream = csv.format({ headers: true });
    const writableStream = fs.createWriteStream(path.resolve(__dirname, 'exports', filename));
    
    csvStream.pipe(writableStream);
    
    applications.forEach(app => {
      // CSVに書き込むデータ整形
      const csvData = {
        '申込ID': app.id,
        '申込日時': app.createdAt.toLocaleString('ja-JP'),
        'ステータス': app.status === 'completed' ? '完了' : '保留中',
        '申込種別': app.applicationType === 'individual' ? '個人' : '法人',
        '氏名/担当者名': app.applicationType === 'individual' ? app.fullName : app.contactPerson,
        'フリガナ': app.applicationType === 'individual' ? app.furigana : '',
        '会社名': app.applicationType === 'corporate' ? app.companyName : '',
        '部署名': app.applicationType === 'corporate' ? (app.department || '') : '',
        'メールアドレス': app.email,
        '電話番号': app.phoneNumber,
        '参加イベント': app.eventType === 'seminar' ? 'セミナー' : app.eventType === 'workshop' ? 'ワークショップ' : 'カンファレンス',
        '参加希望日': app.participationDate,
        '参加人数': app.numberOfPeople === '5' ? app.exactNumber : app.numberOfPeople,
        '備考': app.notes || '',
        '情報取得元': app.hearAbout || ''
      };
      
      csvStream.write(csvData);
    });
    
    csvStream.end();
    
    writableStream.on('finish', () => {
      resolve(filename);
    });
    
    writableStream.on('error', error => {
      reject(error);
    });
  });
}

// 申込データのExcelエクスポート
function exportToExcel() {
  return new Promise((resolve, reject) => {
    const filename = `applications_${Date.now()}.xlsx`;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('申込一覧');
    
    // ヘッダー設定
    worksheet.columns = [
      { header: '申込ID', key: 'id', width: 36 },
      { header: '申込日時', key: 'createdAt', width: 20 },
      { header: 'ステータス', key: 'status', width: 12 },
      { header: '申込種別', key: 'type', width: 10 },
      { header: '氏名/担当者名', key: 'name', width: 20 },
      { header: 'フリガナ', key: 'furigana', width: 20 },
      { header: '会社名', key: 'company', width: 30 },
      { header: '部署名', key: 'department', width: 20 },
      { header: 'メールアドレス', key: 'email', width: 30 },
      { header: '電話番号', key: 'phone', width: 15 },
      { header: '参加イベント', key: 'event', width: 15 },
      { header: '参加希望日', key: 'date', width: 15 },
      { header: '参加人数', key: 'people', width: 10 },
      { header: '備考', key: 'notes', width: 40 },
      { header: '情報取得元', key: 'source', width: 15 }
    ];
    
    // ヘッダースタイル
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0E0E0' }
    };
    
    // データ追加
    applications.forEach(app => {
      worksheet.addRow({
        id: app.id,
        createdAt: app.createdAt.toLocaleString('ja-JP'),
        status: app.status === 'completed' ? '完了' : '保留中',
        type: app.applicationType === 'individual' ? '個人' : '法人',
        name: app.applicationType === 'individual' ? app.fullName : app.contactPerson,
        furigana: app.applicationType === 'individual' ? app.furigana : '',
        company: app.applicationType === 'corporate' ? app.companyName : '',
        department: app.applicationType === 'corporate' ? (app.department || '') : '',
        email: app.email,
        phone: app.phoneNumber,
        event: app.eventType === 'seminar' ? 'セミナー' : app.eventType === 'workshop' ? 'ワークショップ' : 'カンファレンス',
        date: app.participationDate,
        people: app.numberOfPeople === '5' ? app.exactNumber : app.numberOfPeople,
        notes: app.notes || '',
        source: app.hearAbout || ''
      });
    });
    
    // ファイル保存
    workbook.xlsx.writeFile(path.resolve(__dirname, 'exports', filename))
      .then(() => {
        resolve(filename);
      })
      .catch(error => {
        reject(error);
      });
  });
}

// API エンドポイント

// 申込フォーム送信
app.post('/api/submit', applicationValidationRules, (req, res) => {
  // バリデーションエラーの確認
  const errors = validationResult(req);
  
  if (!errors.isEmpty()) {
    return res.status(400).json({ errors: errors.array() });
  }
  
  // 申込データの保存
  const application = saveApplication(req.body);
  
  // 確認メール送信
  sendConfirmationEmail(application);
  
  // 事務局への通知メール送信
  sendNotificationEmail(application);
  
  // 決済ページへのリダイレクトURL生成
  const redirectUrl = `https://決済サービスURL?order_id=${application.id}&amount=1000`;
  
  res.status(200).json({
    success: true,
    message: 'お申し込みを受け付けました',
    application_id: application.id,
    redirect_url: redirectUrl
  });
});

// 決済完了コールバック
app.post('/api/payment/callback', (req, res) => {
  const { order_id, status } = req.body;
  
  if (status === 'success') {
    // 申込ステータスを「完了」に更新
    const updatedApplication = updateApplicationStatus(order_id, 'completed');
    
    if (updatedApplication) {
      res.status(200).json({ success: true });
    } else {
      res.status(404).json({ success: false, message: '申込情報が見つかりません' });
    }
  } else {
    res.status(400).json({ success: false, message: '決済に失敗しました' });
  }
});

// 申込データのCSVエクスポート
app.get('/api/export/csv', async (req, res) => {
  try {
    const filename = await exportToCsv();
    res.download(path.resolve(__dirname, 'exports', filename), filename, (err) => {
      if (err) {
        console.error('Download error:', err);
      }
      
      // ダウンロード後にファイルを削除
      fs.unlinkSync(path.resolve(__dirname, 'exports', filename));
    });
  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ success: false, message: 'エクスポートに失敗しました' });
  }
});

// 申込データのExcelエクスポート
app.get('/api/export/excel', async (req, res) => {
  try {
    const filename = await exportToExcel();
    res.download(path.resolve(__dirname, 'exports', filename), filename, (err) => {
      if (err) {
        console.error('Download error:', err);
      }
      
      // ダウンロード後にファイルを削除
      fs.unlinkSync(path.resolve(__dirname, 'exports', filename));
    });
  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ success: false, message: 'エクスポートに失敗しました' });
  }
});

// サーバー起動
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
  
  // エクスポートディレクトリの作成
  const exportDir = path.resolve(__dirname, 'exports');
  if (!fs.existsSync(exportDir)) {
    fs.mkdirSync(exportDir);
  }
});