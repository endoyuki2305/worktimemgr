# worktimemgr
社員の出退勤を月別にまとめ、給与計算時に利用します。出退勤データは各社員が毎日、スマートフォンあるいはPCからSlackの自社ワークスペース、#timesheets内に呟く。それを[勤怠管理bot - みやもとさん](https://github.com/masuidrive/miyamoto)を通じてGoogle Spreadsheetに書き込む。みやもとさんは残業時間、月別勤務時間集計などの機能が未実装である為、それら機能はユーザーが実装する必要がある。worktimemgr(仮名)は、各社員の出退勤シートから指定月のデータを抽出し、休憩時間タイムシートで定義された休憩時間を引いた勤務時間を算出する。
