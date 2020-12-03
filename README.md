# Poi-JDBC

### 简介

由我自己一个人用Java写的零件库房管理项目

利用Poi技术能够生成两个Excel表然后汇总成一个表

汇总表按照编号进行排序，能够对表进行增删改查操作

将三个表中的数据通过JDBC连接本地数据库导入到对应的表中

并且能对三个表进行增删改查等操作

同时能够在本地生成操作后的xxx表.xlsx

### 两个方向

Poi方面：通过手动输入Excel中的信息生成对应的Excel表，对表能够进行增删改查等四个操作，以及两个表进行汇总。

JDBC方面：连接数据库，将本地的Excel的表导入到数据库对应的Table，实现数据上的存储，另一方面，对数据库直接操作可以实现Excel中的信息更新，并且输出到本地指定的文件路径中

### 系统流程

**主界面**

![](https://picture.hs-vae.com/主界面.png)

**生成零件汇总表**

![](https://picture.hs-vae.com/生成零件汇总表修改.png)

**Poi操作Excel表中的零件信息**

![](https://picture.hs-vae.com/Poi操作Excel表中的零件信息.png)

**JDBC操作Excel表中的零件信息**

![](https://picture.hs-vae.com/JDBC连接数据库操作excel表中的零件信息.png)