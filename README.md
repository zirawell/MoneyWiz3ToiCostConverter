# MoenyWiz3 数据迁移 iCost
本项目旨在迁移 MoenyWiz3 的记账数据至 iCost

尽管 iCost 支持很多软件的数据直接导入，可惜暂时还不支持 MoneyWiz3

因此花了半天自己写了这个项目来做数据迁移

## 项目环境配置
本项目使用 `Java` 编写，JDK版本：1.8，需要导入 `Maven` 依赖
```
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.3</version>
</dependency>
<dependency>
    <groupId>com.opencsv</groupId>
    <artifactId>opencsv</artifactId>
    <version>5.7.1</version>
</dependency>
```

## 操作步骤
1. MoneyWiz3 数据导出，文件名为 `Moneywiz.csv`
2. 将项目中的 `originDataPath` 变量赋值步骤1中导出的文件路径
3. 将项目中的 `outputDataPath` 变量赋值你想要输出的文件夹
4. 运行程序，输出文件夹下可能会有多个文件，因为 iCost 导入文件大小限制，本程序做了每5000条数据切分一次文件
5. 依次导入输出文件夹下生成的 iCost 数据

## 注意
导入完成后，可能会出现某些账户的金额和原本在 MoneyWiz3 中的数值对应不上。

这个问题是因为 MoneyWiz3 中创建账户时设置了初始值导致的，而 MoneyWiz3 数据导出时不会输出这个初始值。

所以我们需要做的就是手动平账，在余额不对的账户中添加分类为`其他`的收入，时间设置到最早的记账时间之前，
金额设置为需要平账的金额，备注可以设置为`初始值`，这样后续的统计数据都将正确。

如果不考虑统计数据的正确与否，也可以使用调整余额的功能，在数据不对的账户上直接调整余额至正确的值就行。

另外，由于两款软件的收入/支出分类相差较大，建议在导出之前，先设置好分类对应，以免导入后再做调整。