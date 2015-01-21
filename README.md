# ppt_xfr
作成したパワーポイントのスライドを、削除、表示/非表示切り替え、ノート削除をすることができます。

# 動作環境
* OS:Windows
* Microsoft PowerPoint
* .ppt/.pptx

Windows7 + PowerPoint 2010で動作確認をしています。

ただし`.ppt`の場合、後述するように`/o`オプションを明示する必要があります。

# 使い方
```
> git clone https://github.com/MATSUMOTO-Takashi/ppt_xfr.git
> cd ppt_xfr
> main.bat /f:config.json /o:output.pptx test.pptx
```

実行時の書式：  
`main.bat [/f] [/o] 元ファイル`

# 実行時オプション
実行時に以下のオプションを指定することができます。

## /f
指定したコンフィグファイルを使用します。

省略した場合`config.json`を読み込みます。

## /o
出力ファイルのファイル名を指定します。

省略した場合`output.pptx`を出力します。

**注意！**：元ファイルが2007以前の`.ppt`の場合、明示的に`/o`を指定しないとエラーとなります。

## /h
ヘルプを表示します。

# 変換用設定ファイル
ex) config_example.json

```
{
  "slide": {
    "show": -1,
    "hidden": [],
    "del": [3, 5, 24]
  },
  "note": {
    "del": -1
  }
}
```

## slide.show
スライドショー時の表示設定です。

type: -1 or array

-1を設定すると全てのスライドに対して表示設定を行います。

指定したスライドのみ表示設定を行いたい場合は`[1, 2, 3]`のように配列型を使用し、スライド番号を羅列してください。

## slide.hidden
スライドショー時の非表示設定です。

type: array

指定したスライドの非表示設定をします。`[1, 2, 3]`のように配列型を使用し、スライド番号を羅列してください。

slide.showとslide.hiddenで同じ番号が指定された場合、**表示状態**となります。

## slide.del
スライドの削除をします。

type: array

指定したスライドを削除します。`[1, 2, 3]`のように配列型を使用し、スライド番号を羅列してください。

## note.del
ノートの削除をします。

type: -1 or array

-1を設定すると全てのスライドのノートを削除します。

指定したスライドのノートのみを削除したい場合は`[1, 2, 3]`のように配列型を使用し、スライド番号を羅列してください。

## スライド番号の指定について
スライドの削除は一連の操作の最後に実行されます。

よって、スライドの番号はパワーポイントを開いた際に振られているスライド番号をそのまま指定してください。
