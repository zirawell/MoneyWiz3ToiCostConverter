package com.zlog;


import com.opencsv.CSVReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.ParseException;
import java.util.*;

/**
 * @Description:
 * @Author: Dark Wang
 * @Create: 2024-10-20 09:43
 **/
public class MoneyWiz3ToiCostConverter {
    // 源文件路径
    private static final String originDataPath = "/Users/zirawell/Downloads/Moneywiz.csv";
    // 输出文件路径
    private static final String outputDataPath = "/Users/zirawell/Downloads/icost_data/";
    // 输出文件表头
    private static final String[] HEADERS = {
            "日期", "类型", "金额", "一级分类", "二级分类", "账户1", "账户2", "备注", "货币", "标签"
    };

    public static void main(String[] args) {
        List<RowData> inputList = getExcelData();
        int total = inputList.size();
        int loopSize = total / 5000;
        // 输出全部数据
        outPutData(inputList,"total");
        // 为导入方便，每隔5000条输出一个文件
        for (int i = 0; i < loopSize; i++) {
            outPutData(inputList.subList(i * 5000, (i + 1) * 5000), i + "");
        }
        outPutData(inputList.subList(loopSize * 5000,total), loopSize + "");
        System.out.println("processed " + total + " rows");

    }

    public static void outPutData(List<RowData> inputList, String count) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("iCost Data");
        // 创建表头行
        Row headerRow = sheet.createRow(0);
        // 设置表头内容
        for (int i = 0; i < HEADERS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(HEADERS[i]);
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            style.setFont(font);
            cell.setCellStyle(style);
        }
        int rowNum = 1;
        for (int i = 0; i < inputList.size(); i++) {
            RowData rowData = inputList.get(i);
            if ("转账".equals(rowData.getType()) && "".equals(rowData.getAmount())) {
                continue;
            }
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(rowData.getDate());
            row.createCell(1).setCellValue(rowData.getType());
            row.createCell(2).setCellValue(rowData.getAmount());
            row.createCell(3).setCellValue(rowData.getFirstCat());
            row.createCell(4).setCellValue(rowData.getSecondCat());
            row.createCell(5).setCellValue(rowData.getAccount1());
            row.createCell(6).setCellValue(rowData.getAccount2());
            row.createCell(7).setCellValue(rowData.getComment());
            row.createCell(8).setCellValue(rowData.getCurrency());
            row.createCell(9).setCellValue(rowData.getTag());
        }
        // 调整列宽
        for (int i = 0; i < HEADERS.length; i++) {
            // 自动调整列宽
            sheet.autoSizeColumn(i);
        }
        File file = new File(outputDataPath + "data_" + count + ".xlsx");

        try {
            // 如果文件的父目录不存在，创建父目录
            if (!file.getParentFile().exists()) {
                file.getParentFile().mkdirs();
            }

            // 创建文件
            if (file.createNewFile()) {
                System.out.println("文件创建成功：" + file.getPath());
                System.out.println("rowNum: " + rowNum);
            } else {
                System.out.println("文件已存在：" + file.getPath());
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        // 写入文件
        try (FileOutputStream fos = new FileOutputStream(outputDataPath + "data_" + count + ".xlsx")) {
            // 将数据写入Excel文件
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                // 关闭工作簿
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    public static List<RowData> getExcelData() {
        List<RowData> resultList = new ArrayList<>();
        try (CSVReader reader = new CSVReader(
                new InputStreamReader(Files.newInputStream(Paths.get(originDataPath)), StandardCharsets.UTF_16))) {
            // 读取所有行
            List<String[]> records = reader.readAll();
            /* record 格式：命名,"当前余额","账户","转账","描述","交易对象","分类","日期","时间","收入","费用","货币","支票号码"
             * record[0]-命名
             * record[1]-当前余额
             * record[2]-账户
             * record[3]-转账
             * record[4]-描述
             * record[5]-交易对象
             * record[6]-分类
             * record[7]-日期
             * record[8]-时间
             * record[9]-收入
             * record[10]-费用
             * record[11]-货币
             * record[12]-支票号码
             */
            for (String[] record : records) {
                RowData data = null;
                if (record.length == 13) {
                    // 去除非法数据
                    if ("命名".equals(record[0])) {
                        continue;
                    }
                    int count = 0;
                    for (int j = 4; j < record.length; j++) {
                        if ("".equals(record[j])) {
                            count++;
                        }
                    }
                    if (count == 9) {
                        continue;
                    }
                    data = dataConvert(record);
                    resultList.add(data);
                    for (String field : record) {
                        System.out.print(field + ",");
                    }
                }
                System.out.println();

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return resultList;
    }

    private static RowData dataConvert(String[] record) throws ParseException {
        record = preProcessData(record);
        RowData data = new RowData();
        // 日期处理
        data.setDate(formatDateStr(record[7], record[8]));
        // 货币处理
        data.setCurrency(record[11]);
        // 描述处理
        data.setComment(record[4]);
        // 账户1处理
        data.setAccount1(record[2]);
        // 分类处理
        if (record[6].contains(">")) {
            String[] cats = record[6].split(">");
            data.setFirstCat(cats[0]);
            data.setSecondCat(cats[1]);
        } else {
            data.setFirstCat(record[6]);
        }
        // 交易对象->标签
        if (!record[5].isEmpty()) {
            data.setTag("#" + record[5]);
        }

        // 类型处理
        String type;
        if (!"".equals(record[2]) && !"".equals(record[3])) {
            type = "转账";
        } else {
            type = record[10].contains("-") ? "支出" : "收入";
        }
        data.setType(type);
        // 转账处理 - 需加上账户2
        if ("转账".equals(type)) {
            data.setAccount2(record[3]);
        }
        // 金额处理
        if ("收入".equals(type)) {
            data.setAmount(record[9].replace("-", "").replace(",", ""));
        } else {
            data.setAmount(record[10].replace("-", "").replace(",", ""));
        }
        return data;
    }

    /**
     * 特殊字符预处理
     *
     * @param record
     * @return
     */
    private static String[] preProcessData(String[] record) {
        String[] data = Arrays.copyOf(record, record.length);
        for (int i = 0; i < record.length; i++) {
            if (record[i].contains(" ")) {
                data[i] = data[i].replace(" ", "");
            }
        }
        return data;
    }


    static class RowData {
        //日期-格式: 2011年01月11日 12:00:00
        String date;
        //类型-格式: 支出/收入/转账
        String type;
        //金额
        String amount;
        //一级分类
        String firstCat;
        //二级分类
        String secondCat;
        //账户1
        String account1;
        //账户2
        String account2;
        //备注
        String comment;
        //货币
        String currency;
        //标签
        String tag;

        public String getDate() {
            return date;
        }

        public void setDate(String date) {
            this.date = date;
        }

        public String getType() {
            return type;
        }

        public void setType(String type) {
            this.type = type;
        }

        public String getAmount() {
            return amount;
        }

        public void setAmount(String amount) {
            this.amount = amount;
        }

        public String getFirstCat() {
            return firstCat;
        }

        public void setFirstCat(String firstCat) {
            this.firstCat = firstCat;
        }

        public String getSecondCat() {
            return secondCat;
        }

        public void setSecondCat(String secondCat) {
            this.secondCat = secondCat;
        }

        public String getAccount1() {
            return account1;
        }

        public void setAccount1(String account1) {
            this.account1 = account1;
        }

        public String getAccount2() {
            return account2;
        }

        public void setAccount2(String account2) {
            this.account2 = account2;
        }

        public String getComment() {
            return comment;
        }

        public void setComment(String comment) {
            this.comment = comment;
        }

        public String getCurrency() {
            return currency;
        }

        public void setCurrency(String currency) {
            this.currency = currency;
        }

        public String getTag() {
            return tag;
        }

        public void setTag(String tag) {
            this.tag = tag;
        }
    }

    /**
     * 注意此处MoneyWiz3默认导出时间格式为YYYY/DD/MM
     * 根据需要进行更改
     * iCost的日期格式为 2011年01月11日 12:00:00
     *
     * @param date 日期
     * @param time 时间
     * @return String
     */
    public static String formatDateStr(String date, String time) {
        String year = null;
        String month = null;
        String day = null;
        String[] arr = date.split("/");
        year = arr[0];
        month = arr[2];
        day = arr[1];
        return year + "年" + month + "月" + day + "日" + " " + time;
    }

}

