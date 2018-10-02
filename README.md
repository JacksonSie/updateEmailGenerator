# updateEmailGenerator

- 這個工具可以用來產生每周更新通報 EDM ，省去每次產出電子報都要重新排版的困擾。
- 藉由同時產出 html 與 doc 的方式，同時產生出 outlook 與 mail2000 不會跑版的檔案。

## 使用方式:
```sh
1. 更新 test/email.xlsx ，並依據updatesTitle排序 (更新方式請看email.xlsx)
2. $ python test/generator.py
3. 在產出 test/email.html 與 test/email.doc ，並依據需要微調排版
```

-----

## changlog
> Date:   Mon Jul 25 19:32:03 2016 +0800
+ 目前最後手工3步將email完成: 
    > 1. 全選 ->右鍵 -> 段落 ->  與前段間距0,與後段間距0,行距1
    > 1. 全選 -> 更改字體為 times new roman -> 標楷體 -> 項目符號 windings
    > 1. 判斷 suggestion 有 "廠商已發布更新" and "已有攻擊手法被公開" -> 紅色標記

> Date:   Mon Jul 25 10:53:24 2016 +0800
+ EDM想兼容 html 跟 outlook 基本上是很難(詳見 reference)，所以更動 3 項 
    > 1. 整理code，包括 define main ， vim set tabstop=4
    > 1. 另外多存一個 word
    > 1. import 測試
> + ref. http://blog.brain1981.com/325.html
