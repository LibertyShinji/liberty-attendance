# 出勤・有給管理システム 導入手順

LINE（LIFF）＋ Google Apps Script（GAS）＋ Google スプレッドシートで動く、社員10〜15名規模向けの勤怠システムです。
サーバー費用ゼロ（GAS無料枠＋GitHub Pages＋LINE公式アカウント）で運用できます。

> このZIPは実際に中古車販売会社で本番稼働しているシステム（2026-07-30時点の最新版）から、会社固有の情報
> （ID・氏名・URL）をプレースホルダーに置き換えたものです。`YOUR_XXX` と書かれた箇所を自分の環境の値に置き換えてください。
> **Claude Code / Claude に丸ごと渡して「これを土台にセットアップして」と頼む前提で書いています。**
> 開発を続ける場合は `開発ガイド_Claude用.md` を、
> **自社に合わせて機能を変えたい／削りたい場合は `カスタマイズガイド.md`** を必ず読ませてください。
> （全員正社員でパートがいない、LINE通知は要らない、公休の上限日数が違う——そういう会社でも
> カスタマイズガイドの手順どおりに調整すればそのまま使える構成です）

## できること

| 画面（LIFF） | 機能 |
|---|---|
| 打刻 | 出勤・退勤のワンタップ打刻（GPS記録・時間外/休日打刻の管理者通知）＋「今日は急に休みます」ボタン |
| 有給申請 | 全休/午前半休/午後半休（0.5日）→ 管理者承認 → 残日数自動減算・年度（4月始まり）管理 |
| シフト希望申告 | 出勤できる日・時間帯をタップ＋時刻入力で申告 → 管理者が確定（予約対応のシフト制向け・2026-08追加） |
| 休み希望 | 公休（無給の定例休みの月上限日数管理・正社員6日/パート15日）の希望日をタップ申請＋確定公休を「出勤に戻す」申請。シフトの有無とは独立して運用 |
| 自分の休み | 本人カレンダー（出勤予定/出勤済/欠勤/公休/有給/シフト無しの色分け・シフト確定日は時間帯表示）＋パートは給与目安カード（今月ここまで/月末見込/手取り目安/月額目標） |
| 全員カレンダー | 日別の出勤人数と誰が休み・確定シフト時間帯かの一覧 |
| 管理者 | 出勤状況・月次一覧（付け忘れ可視化）・シフト確定・公休確定（承認済み有給も緑表示）・有給承認・残日数（名前タップで消化日一覧）・給与集計・代理打刻修正・社員マスタ管理 |

自動機能：出勤忘れ通知（各自の確定シフト開始時刻から60分間）／退勤忘れ通知（19-22時・Libertyでは送信停止／後述）／前日以前の退勤もれを役員へ通知／月次集計CSVメール。
※シフトが確定していない日は「非勤務日」として扱われ、欠勤判定・出勤忘れ通知の対象になりません。

### Liberty運用メモ：退勤忘れの本人向け通知は使わない（2026-08〜）
退勤忘れは吉田さんが管理画面（「前日以前の退勤もれ」役員通知＋月次一覧の「付け忘れ」表示）で見つけて手動修正する運用とし、
本人への自動DM（19-22時のリマインド）は送らない方針です。コード変更は不要で、**セットアップ時に `COREMINDER_DRYRUN` を
`'false'` に設定しない（未設定のままにする）だけで送信されません**（未設定＝ログのみのドライラン動作がデフォルトのため）。
一方、出勤忘れ通知（本人へ）は使う想定なので、こちらは `CIREMINDER_DRYRUN` を `'false'` に設定して有効化してください。

## 必要なもの

1. Googleアカウント（スプレッドシート＋GAS）
2. LINE公式アカウント（Messaging APIチャネル）— 通知DM用。フリープランは月200通まで（通知が多いなら有料プラン検討）
3. LINEログインチャネル（LIFF用）— **Messaging APIと同じプロバイダーに作ること**（違うとUserIDが別値になり紐付け不能）
4. GitHub アカウント（Pages で画面HTMLを無料ホスティング）

## セットアップ手順（この順番で）

### 1. スプレッドシート作成
新規スプレッドシートを作成し、IDを控える（URLの `/d/` と `/edit` の間）。
シートはGAS側の `initAllSheets()`（無ければ `initSheet` を呼ぶ各関数）が自動作成するので手動で作らなくてよい。

### 2. GASプロジェクト作成
1. script.google.com で新規プロジェクト →「Code.gs」「ShiftManagement.gs」「RichMenuImage.gs」を作成し、このZIPの同名ファイルの中身を貼り付け（claspが使えるなら `clasp push` でも可）。
2. HTMLファイル6枚（Attendance/LeaveRequest/AdminDashboard/Calendar/KibouRequest/TeamCalendar）もGASにHTMLファイルとして追加。
   ※実際の画面配信はGitHub Pagesから行う（後述）。GAS内のHTMLは原本管理用。
3. `initProperties()` 内の `YOUR_SPREADSHEET_ID` を書き換えて1回実行 → スクリプトプロパティが入る。
4. `initAllSheets()` を実行してシート雛形を作成。
5. デプロイ →「ウェブアプリ」→ 実行ユーザー=自分／アクセス=全員 → **デプロイURL（/exec）を控える**。

### 3. LINE側の設定
1. Messaging APIチャネルのアクセストークン（長期）を発行 → スクリプトプロパティ `LINE_CHANNEL_ACCESS_TOKEN` に設定。
   **Webhookは設定不要**（全機能LIFF完結。既存Botと同居する場合はWebhookを触らないこと）。
2. LINEログインチャネルを作成（同一プロバイダー！）→「公開済み」にする（開発中のままだと社員が開けず400エラー）。
3. LIFFアプリを7個作成（サイズFull推奨）。エンドポイントURLはひとまず仮でよい（後でGitHub Pagesに変更）：
   - 打刻 / 有給申請 / 管理者 / 自分の休み(カレンダー) / 休み希望(公休) / 全員カレンダー / シフト希望申告
4. 7個のLIFF IDをスクリプトプロパティに設定：
   `LIFF_ATTENDANCE_ID / LIFF_LEAVE_ID / LIFF_ADMIN_ID / LIFF_CALENDAR_ID / LIFF_KIBOU_ID / LIFF_TEAM_ID / LIFF_SHIFT_ID`

### 4. 画面HTMLをGitHub Pagesで公開（重要）
**GASのHtmlServiceはサンドボックスiframe配信のためLIFFのliff.init()が動きません。画面は必ず外部ホスティングに置きます。**
1. `build_static.py` の冒頭の `EXEC_URL`（手順2-5のURL）と6つのLIFF IDを書き換える。
2. `Calendar.html / KibouRequest.html / LeaveRequest.html` 内の `YOUR_KIBOU_LIFF_ID` `YOUR_LEAVE_LIFF_ID` を検索して実IDに置換（画面間リンク用のハードコード箇所）。
3. `python build_static.py` を実行 → `docs/` に6枚の静的HTMLが生成される。
4. GitHubにPublicリポジトリを作り、`docs/` の6ファイルをアップロード → Settings → Pages を有効化（main / root）。
5. LINEコンソールで6つのLIFFのエンドポイントURLを `https://<user>.github.io/<repo>/attendance.html` などに変更。

### 5. 社員登録
`registerInitialEmployees()` のサンプル配列を自社スタッフに書き換えて1回実行。
- 一般社員のLINE UserIDは空でOK → 本人が初回に打刻画面を開くと名前を選んで自動紐付け（bindMyLineId）。
- **管理者（役員）のUserIDだけは必須**（承認通知DMの宛先）。
- 有給の初期残高がある場合は `loadLeave2026()` を書き換えて実行。
- パートがいる場合は社員マスタI列を「パート」に（管理者画面の⚙社員タブからも変更可）。時給・控除は同タブで入力。

### 6. トリガー設定（GASエディタで1回）
- `setupClockOutReminderTrigger()` を実行 → 15分毎トリガーが入る（退勤忘れ・出勤忘れ・退勤もれ役員通知はすべてこの1本から動く）。
- `updateAnnualLeave` を年間タイマー（1月1日 午前0時）で登録（年度更新）。
- メールCSVを使うなら `setupEmailAuth()` を1回実行して権限承認。

### 7. リッチメニュー（任意）
`setupRichMenu6(管理者UserId)` で6ボタンメニューを自動作成（画像は `RichMenuImage.gs` に内蔵のbase64、元画像は `richmenu6_lineart.jpg`）。
自作する場合の条件：幅800-2500px・縦横比1.45以上・1MB未満・JPEG/PNG。

## ⚠️ ハマりどころ（先人の教訓）

- **GASでLIFF画面配信は不可能**（手順4の理由）。画面は必ずGitHub Pages等の外部ホスティング。
- fetchは全て `Content-Type: text/plain;charset=utf-8` にする（CORSプリフライト回避。application/jsonにすると死ぬ）。
- **APIテストで日本語をPOSTするときは必ずASCIIエスケープ**（Pythonなら `json.dumps(..., ensure_ascii=True)`）。Windowsのcurl直叩き＋日本語リテラルは文字化けする。
- GAS doPostは302リダイレクトで結果を返す（curlは2段階、Python urllibは自動追従）。
- 時刻整形は必ず **Asia/Tokyo**（`fmt` ヘルパー使用）。GMTで整形すると-9時間ズレる。appsscript.jsonのtimeZoneも確認。
- 有給承認は `LockService` で排他制御済み（同時承認の2重減算対策）。**スプレッドシートの申請行を手動削除しないこと**（残日数が狂う。取消は管理画面のAPIで）。
- HTMLを修正したら build_static.py 再実行 → docs/ をGitHubへ再アップ。GitHub PagesのCDNキャッシュは数分残る。
- LIFFエンドポイント変更後、社員がLINEを開き直すまで古い画面が出ることがある。

## ファイル構成

| ファイル | 役割 |
|---|---|
| Code.gs | 本体：doPostルーター・打刻・有給・給与・通知・社員管理・セットアップ関数 |
| ShiftManagement.gs | カレンダー・公休（希望/確定/取消希望）・全員カレンダー・シフト（出勤予定の希望/確定） |
| RichMenuImage.gs | リッチメニュー画像（base64内蔵） |
| *.html（7枚） | 画面の原本（GASテンプレート形式） |
| build_static.py | 原本→GitHub Pages用静的HTML変換（URL/LIFF ID注入） |
| appsscript.json | GASマニフェスト（タイムゾーン等） |
| 開発ガイド_Claude用.md | アーキテクチャ・API一覧・データ構造（開発を続けるAI向け） |
| カスタマイズガイド.md | 自社に合わせた調整方法（雇用形態・公休上限・LINE通知の削り方・機能の取捨選択マップ） |
