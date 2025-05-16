# Sql-to-Excel
根据SQL建表语句提取数据表内容写入表格

# 实现准备

首先要准备python环境，这里本地运行采用了python 3.7版本解释器下的虚拟环境，相关依赖库可通过执行

```shell
pip install -r requirements.txt
```

# 输入格式

这里不能保证可以适配所有格式的建表语句，本人没有时间去全部测试，目前已知如下的格式可正常运行:

```
create table 表空间.表名1(
       列名1 word(30) not null,
       列名2 word(30) not null,
       列名3 number(6,2) not null,
);
alter table 表空间.表名1 add primary key (列名1,列名2);
alter table 表空间.表名1 comment on column 列名3 is '测试1';

create table 表空间.表名2(
       列名1 word(30) not null,
       列名2 word(30) not null,
       列名3 number(6,2) not null,
);
alter table 表空间.表名2 add primary key (列名1,列名2);
alter table 表空间.表名2 comment on column 列名3 is '测试1';
```

支持多组输入，不同表建表语句之间要空行，最后Ctrl + D结束输入，同级目录内会出现名为output.xlsx的文件，起初没有该文件会创建，若已有则先清空后写入。

# 效果示例

不同表会存在不同的工作表中，名称与表名一致

![image-20250516140315746](\image\result.png)