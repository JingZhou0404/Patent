Patent
======

写在代码结构之前的话：
代码的作用
上传的代码是根据国资局给的不完全的专利号在专利之星中搜索完全的专利号。具体而言，1994年之前的专利号是不带有验证码的，1994年后国资局
进行改革使用了验证码，由于很多原因(我们也不是很清楚)有大量的专利号没有校验码，但是专利之星网站上的专利号有90%是带有验证码的，此代
码的作用就是对无验证码的专利号在专利之星上进行搜索，获得验证码。例如，无验证码的专利号85109282，有验证码的专利号85109282.9。

选择上传这个代码是因为代码量少，上次您问过我关于这个项目的实现，这样您看着也方便。

这个项目主要遇到3个问题，1. 如何判断网页是否加载完成。 2.如何判断Ajax数据加载完成。 3.如何自动点掉弹出的对话框
1. 判断网页加载完成是利用 Application.DoEvents()来等待网页加载使用布尔变量在webBrowser1_DocumentCompleted中判断页面是否加载完成
2. 利用Timer等待来解决这个问题
3. 在webBrowser1_DocumentCompleted 和 webBrowser1_Navigated中运行js。

代码结构：

这个项目当时由于是处理了多个东西，所以有些多余的页面。
真正有用的文件是Patent / PatentData / PatentData / patentStar.cs 以及 Patent / PatentData / PatentData / SQLHelper.cs
SQLHelper.cs用来处理数据库的，patentStar.cs用来爬取数据的，其它文件都是处理别的问题的。

曾经和一个工作过的大哥一起做过一个食疗网站，网站上线了，并有用户使用，链接为http://www.mxfslys.com/
