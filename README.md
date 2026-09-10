TOEIC Demo

TOEIC の Listening / Reading 模擬試験を自動採点する Python ツールです。

Excel に回答を入力し、正解シートと比較して以下を自動算出します。

正答数
誤答数
正答率
不正解問題一覧

採点結果は Excel の結果シートへ記録され、実行ログは output.log に出力されます。

flowchart TD

    A["問題用紙.xlsx<br/>回答用紙シート"] --> D[main.py]

    B["解答用紙.xlsx<br/>L_1～L_5<br/>R_1～R_5"] --> D

    D --> E[採点処理]

    E --> F[結果シートへ保存]
    E --> G[output.logへ出力]



toeic_demo
├── LICENSE
├── README.md
├── output.log
├── 問題用紙.xlsx
├── 解答用紙.xlsx
└── src
    └── main.py



