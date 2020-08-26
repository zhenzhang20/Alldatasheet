---

---

 

最近需要从[全球电子元器件数据手册库](https://www.alldatasheet.com/) 网上下载一些datasheet文件，手动下载比较耗时，写个下载脚本可以节省很多时间，且可以在需要下载内容更新后重复使用，提高效率。

### 涉及到的知识点

1. Python读写Excel
2. Python requests 模块

### Excel管理表格

![datasheet_excel](https://github.com/zhenzhang20/zhenzhang20.github.io/tree/master/2020/08/25/2020-08-25-Use-Python-Requests-Download-Alldatasheet/2020-08-25-datasheet_excel.png)



### 需求

1. 根据Excel管理表格中“链接”列内容，下载产品的pdf格式datasheet文件
2. 下载文件重命名为“型号规格.pdf”
3. 下载文件存储到“分类”列指定的文件夹中
4. 下载到pdf文件后，更新“下载”列的No为Yes



### 运行过程

在命令行运行结果示例：

```
C:\Users\Administrator\Download>python download.py
['A', 'ADG201AKR', 'No', 'https://www.alldatasheet.com/datasheet-pdf/pdf/48648/AD/ADG201AKR.html']
尝试URL连接中...
获取返回内容信息文本，如果是pdf，比较耗时，请等待...
获取返回内容信息文本完成。
尝试存储文件：ADG201AKR.pdf
存储文件：ADG201AKR.pdf完成
......
['C', 'TPS5450DDA', 'No', 'https://www.alldatasheet.com/datasheet-pdf/pdf/180875/TI/TPS5450DDA.html']
尝试URL连接中...
获取返回内容信息文本，如果是pdf，比较耗时，请等待...
获取返回内容信息文本完成。
下载太频繁，等待五分钟......，等待中
尝试URL连接中...
获取返回内容信息文本，如果是pdf，比较耗时，请等待...
获取返回内容信息文本完成。
尝试存储文件：TPS5450DDA.pdf
存储文件：TPS5450DDA.pdf完成
['C', 'DCP010512BP', 'No', 'https://www.alldatasheet.com/datasheet-pdf/pdf/527520/TI1/DCP010512BP-U.html']
尝试URL连接中...
获取返回内容信息文本，如果是pdf，比较耗时，请等待...
获取返回内容信息文本完成。
尝试存储文件：DCP010512BP.pdf
存储文件：DCP010512BP.pdf完成
```

### 说明

1. 连续下载10个pdf后，网站会提示输入验证码，网页中含有信息“Download is temporarily unavailable”。观察发现，在遇到提示后，等待几分钟，便可以继续下载。程序中：在遇到下载失败后，会等待了5分钟，然后继续下载。
2. 本程序唯一难点是在分析网页时，要能找到post的参数：

```
        params = {
            "tmpinfo1aa": "abc",
        }
```



查找位置如下： 

![alldatasheet_download_post](https://github.com/zhenzhang20/zhenzhang20.github.io/tree/master/2020/08/25/2020-08-25-Use-Python-Requests-Download-Alldatasheet/2020-08-25-alldatasheet_download_post.png)



[View Blog](http://www.xiejiashan8.com/2020/08/25/2020-08-25-Use-Python-Requests-Download-Alldatasheet/)
[Fork/Star on Github](https://github.com/zhenzhang20/Alldatasheet)


