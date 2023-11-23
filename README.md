# easyGet_arxivPaper

## 介绍
本仓库提供了一种简单的方法可以迅速的获得在既定检索关键词下的文献列表并且输出成一个便于阅读的word文档

## 需要安装的python库

```powershell
pip install feedparser
pip install numpy
pip install pandas
pip install datetime
pip install feedparser
pip install python-docx
```

## 使用说明

本程序需要用户手动指定搜索条件，搜索条件的语法规则如下：

| prefix | explanation           |
| ------ | --------------------- |
| ti     | 标题                  |
| au     | 作者                  |
| abs    | 摘要                  |
| co     | 会议                  |
| jr     | JournalReference      |
| cat    | SubjectCategory       |
| rn     | ReportNumber          |
| id     | Id(useid_listinstead) |
| all    | Alloftheabove         |

同时用户可以指定输出的文档名称，参数明伟save_name。
