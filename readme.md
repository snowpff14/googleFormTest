# 使い方
## 導入
* pythonをインストール(3.7で確認)
* 以下のコマンドでライブラリのインストールを行う
```
pip install numpy
pip install selenium
pip install pandas
pip install openpyxl
pip install xlrd
```
＊OSはwindows10似て確認
* ブラウザのインストール
    * __FireFox__ を使用するためFireFoxがインストールされていなければインストールすること。
        * ほかのブラウザでもdriverの指定を変えれば動くと思われるが・・・
* webdriverの配置
    * [こちら](https://github.com/mozilla/geckodriver/releases)からダウンロード後解凍して`C:\webdriver` に格納する。配置場所を変えた場合はソースの該当箇所を変更する。
    Chromeの場合は[こちら](http://chromedriver.chromium.org/downloads)から入手

## 使い方

1. `data/フォーム.xlsx`に各入力内容に応じて値を入れる。
1. `googleFormReport`を実行する。
1. 1行分終わるごとに確認のポップアップが表示されるので応答することで次の行が実行される。


## ソースの説明

* 以前公開したクラスを継承して使用している。
* アラートダイアログに関しては基底クラスに定義していませんが定義するかも・・・
* エクセルなどでまとめてデータがある場合一気に自動入力できるような処理のサンプルとして作成。
    * 実際にはこの機能を応用して勤務表の入力やとあるシステムの業務支援として使っている。
    * 通常の方法では操作できないので、今回はGoogleFormのプルダウンを操作できるように対応。
        * プルダウンの `▽` の部分をクリックして要素を選択するようにしている。
* logの設定など見よう見まねでやっているのでベストプラクティスではないかもしれない

