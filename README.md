# 複数のスプレッドシートのデータを集約する
特定のディレクトリ以下にあるスプレッドシートのデータを集約する

以下の記事のソースコードまとめです．  
[複数の Google スプレッドシートのデータを集約する](http://qiita.com/kz_takatsu/items/a89e89a4c5e82414ae3f)

* aggregateData.gs ... データを集約するスクリプト本体
* Sidebar.html     ... サイドバーの設定
* Stylesheet.html  ... サイドバーのスタイルシート
* SidebarJS.html   ... ボタンアクションの設定

ディレクトリ構造は以下で行いました．  

```
適当なディレクトリ/  
 ┣ 経理精算書 （ディレクトリ名は集計テスト内の集計設定シートに準ずる）/  
 ┃       ┣ 集約対象のスプレッドシート1  
 ┃       ┣ 集約対象のスプレッドシート2  
 ┃       ...  
 ┗ 集計テスト（集約結果が適用されるスプレッドシート)，ここにaggregateData.gsのコードをAttach  
```


