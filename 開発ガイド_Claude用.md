# 開発ガイド（このシステムを引き継ぐAI/開発者向け）

このドキュメントは、Claude Code等のAIがこのコードベースを土台に開発を続けるための内部仕様書です。
セットアップは `README_導入手順.md` を先に読むこと。

## アーキテクチャ

```
[社員のLINE] → リッチメニュー → LIFF(6画面)
                                   │ 画面本体は GitHub Pages の静的HTML
                                   │ (GASのHtmlServiceはLIFF不可のため)
                                   ▼ fetch POST (text/plain)
                            [GAS ウェブアプリ doPost]  ←— 15分毎トリガー(リマインド類)
                                   │
                                   ▼
                        [Google スプレッドシート(全データ)]
                                   │
                                   ▼ Messaging API push
                        [LINE DM通知（本人リマインド・管理者承認依頼）]
```

- 認証：LIFFの `liff.getProfile().userId` で本人特定（PINなし）。社員マスタのLINE UserID列と突合。
- 画面→GAS間は全て `{type: 'ルート名', userId, ...}` のJSON POST。doPost内のswitchが唯一のルーター。
- 管理者判定は社員マスタの管理者列（isAdmin）。役員=管理者は打刻対象外。

## スプレッドシート構造（シート名はCode.gsのSHEETS定数）

| シート | 列 |
|---|---|
| 社員マスタ | A:社員ID B:氏名 C:PIN(廃止・空) D:入社日 E:年間有給日数 F:LINE UserID G:管理者 H:有効 I:雇用形態(正社員/パート) J:平日時給 K:土曜時給 L:日祝時給 M:控除率% N:月額目標 O:固定控除額 |
| 打刻記録 | A:日付 B:社員ID C:氏名 D:出勤 E:退勤 F:勤務時間 G:出勤場所 H:退勤場所 I:備考 J:区分(平日/土曜/日祝) |
| 有給申請 | A:申請ID B:社員ID C:氏名 D:開始日 E:終了日 F:日数 G:理由 H:ステータス I:申請日時 J:処理日時 K:承認者 L:半休区分 |
| 有給残日数 | A:社員ID B:氏名 C:年度 D:付与合計 E:使用 F:残 |
| 公休設定 | A:社員ID B:日付 C:氏名 D:登録日時 E:ステータス(確定/希望/取消希望) F:半休区分(全休/午前/午後) |
| 急な休み申請 | A:日時 B:日付 C:社員ID D:氏名 E:理由 |
| 有給変更履歴 | 申請/承認/否認/取消/手動訂正の追記専用監査ログ |
| 退勤通知ログ・出勤通知ログ | リマインド送信記録（スロットル制御兼用） |

## 休みのデータモデル（最重要）

休みの入口は3つに統一されている（2026-07に「シフト変更申請」を廃止して整理済み）：

1. **休み希望（公休）** … 無給の定例休み。上限=正社員6日/パート15日/月、半休0.5。
   - 社員が kibou.html でタップ申請 → 公休設定シートに `希望` 行
   - **確定公休をタップ→「出勤に戻す」申請** → `取消希望` 行（確定行はそのまま残す＝承認まで休み扱い）
   - 管理者が admin.html 公休タブ「希望を反映」（希望→追加・取消希望→削除をまとめて取込）→「確定保存」
   - **確定保存は当該社員×当月の全行を消して確定行を書き直す置換方式**（希望・取消希望行もここで消える）
   - **休み希望は明日以降のみ**。当日は下記の急な休みに一本化。
2. **急な休み（当日）** … 打刻画面の控えめリンク。急な休み申請シートに記録＋出勤予定者と役員全員へDM＋
   **公休設定に当日の`希望`行を自動起票**（autoRegisterSuddenKoukyuu_）→管理者が公休タブで確定。
   ※このためsubmitKibouの置換削除は「明日以降の行のみ」対象（当日の自動起票行を消さないため）。
3. **有給申請** … 残日数連動・給与あり。leave.html→管理者の有給タブで承認。年度(4月始まり)・繰越上限20日・
   半休0.5日・LockServiceで2重承認防止・有給変更履歴に全イベント記録。

**読み取り側の鉄則：公休は `ステータス==='確定'` の行だけを休みとして扱う。**
（月次一覧・CSV・給与見込・出勤/退勤リマインド・チームカレンダー全てこの規約。
`取消希望` は同日に確定行が併存するので、本人カレンダー(getCalendarData)だけは
「取消希望行で確定行を上書きしない」特別処理をしている＝day.koukyuuCancelフラグ）

## カレンダーのステータスモデル（calcDayStatus）

`attended / clocked_in / leave / leave_half / koukyuu / koukyuu_half / koukyuu_request / koukyuu_request_half / scheduled(未来) / absent(過去で打刻・公休・有給なし)`
- 半休（有給・公休とも）は半日勤務するので打刻があれば出勤扱い＋バッジ表示。
- **公休を登録していない休業日は absent（欠勤）に見える**仕様。全社休業日も公休登録する運用が前提。

## 主なAPIルート（doPost switch。全て POST text/plain・JSON）

打刻系: liff_clock_in / liff_clock_out / get_today_status / submit_sudden_absence
有給系: submit_leave / get_my_balance / approve_leave / reject_leave / delete_leave / adjust_leave_used / get_leave_log / get_employee_leave（残日数タブの名前タップ用・今年度承認済み一覧・同一期間の重複行は1件に集約）
公休系: get_my_kibou / submit_kibou(dates=[{d,h}], cancels=[d]) / get_koukyuu_admin（承認済み有給`leave`も返す） / save_koukyuu / request_kibou_all
カレンダー: get_calendar / get_team_calendar / get_month_overview
給与系: get_payroll / get_my_payroll / set_my_target / email_monthly_summary
社員管理: get_employees_full / add_employee / retire_employee / reactivate_employee / get_unbound_employees / bind_my_line_id / save_employee_settings
その他: get_dashboard / get_attendance_admin / save_attendance_admin / get_reminder_logs / check_admins_reachable / setup_richmenu6

## 自動通知（1本の15分毎トリガー checkClockOutReminders から全て動く）

- 退勤忘れ：19:00-22:00、出勤済&退勤なしの本人へ約30分毎DM（退勤通知ログでスロットル）
- 出勤忘れ：10:00-11:00、未打刻の本人へ約15分毎DM。除外=確定公休(全休/午前)/承認有給(同)/急な休み登録者。管理者へ未打刻一覧を1日1回
- 前日以前の退勤打刻もれ：朝9時以降1日1回、役員のみへまとめDM（過去14日分）
- ドライラン用ScriptProperty: COREMINDER_DRYRUN / CIREMINDER_DRYRUN（'false'で本番送信）

## パート給与計算（computePartPayroll_）

- 現在給与 = Σ(区分別勤務時間 × 区分別時給)。区分は打刻時に自動記録（祝日はJP_HOLIDAYS_YYYY定数——**年が変わったら追加必要**）
- 月末見込 = 現在 ÷ 出勤日数 × 予定出勤日数（公休/承認有給以外の全日・半休0.5）
- 手取り目安 = 額面 − 固定控除額(社保) − 額面×控除率(雇用保険等)。源泉は未算入（あくまで目安表示）
- 月額目標は本人がカレンダー画面で設定（set_my_target）→「あと◯日で達成」表示

## 開発時の作法・罠（実運用で踏んだもの）

1. **fetchのContent-Typeは text/plain 固定**（CORSプリフライト回避）。
2. **APIテストは ensure_ascii=True のJSONで**（Windows環境から日本語を生POSTすると化ける）。
3. doPostの結果は302リダイレクト先にJSON（Python urllibは自動追従）。
4. 日時整形は `fmt(date,'yyyy/MM/dd')`（Asia/Tokyo）。文字列比較で日付比較する規約（'2026/07/29' > '2026/07/28'）。
5. シートから読んだ日付セルは必ず `fmt(new Date(r[n]),'yyyy/MM/dd')` で正規化してから比較。
6. 有給の残日数を直接いじらない。訂正は adjust_leave_used。申請行の手動削除は絶対禁止（監査ログとfindUnresolvedApproval_が2重申請をブロックする）。
7. 通知を送るAPI（submit系）はテストすると実DMが飛ぶ。テストは遠い将来月＋テスト後にsave_koukyuu空保存等で必ず後始末。バリデーションエラー系は通知前にreturnするので安全にテスト可。
8. HTML修正→build_static.py→docs/をGitHub Pagesへ。GAS側(.gs)修正→clasp push→**既存デプロイIDに対して** `clasp deploy -i <ID>`（新規デプロイにするとURLが変わり全LIFFが死ぬ）。
9. 半休は「有給の半休（0.5日消化・給与あり）」と「公休の半休（枠0.5日・無給）」の2種類。混同注意。
10. 社員追加は管理者画面⚙社員タブ→本人がLIFFを開いて名前を選ぶと自動紐付け（getUnboundEmployees/bindMyLineId）。

## 管理者画面の見どころ（2026-07-30版で追加された表示）

- 公休タブ：本人の休み希望（紫枠）・出勤に戻す希望（橙枠「戻」）に加え、**承認済み有給が緑セル+「有」バッジ**で表示される
  （タップ不可。公休と別枠であることがメモ行にも出る）。データはget_koukyuu_adminの`leave`。
- 残日数タブ：**名前をタップするとその社員の今年度の有給消化日一覧が展開**される（get_employee_leave・lvCacheで再取得なし）。
  一覧合計と残日数シートの「使用」に差がある場合は「※ほかにシステム導入前の使用X日」と注記される
  （導入前にExcel等で消化した分は申請行が存在しないため）。

## カスタマイズ

雇用形態（正社員/パート）・公休上限・LINE通知の要否・祝日・年度などを自社に合わせて変える方法は
**`カスタマイズガイド.md`** にまとめてある。機能を削る改修を頼まれたら、先にそちらの「機能マップ」を読むこと。

## 拡張のヒント（元の会社でやった/検討した順）

- 月次一覧のセルタップ→代理修正はadmin.htmlに実装済み。付け忘れ運用はこれで回る。
- リマインド回数はCOREMINDER/CIREMINDER定数で調整可（LINE無料枠200通/月に注意。人数×回数で意外と消費する）。
- 祝日テーブル（JP_HOLIDAYS_YYYY）は毎年12月に翌年分を追加。
