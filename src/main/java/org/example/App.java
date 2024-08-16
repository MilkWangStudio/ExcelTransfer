package org.example;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.io.resource.ResourceUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.compress.utils.Lists;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * Hello world!
 */
@Slf4j
public class App {
    public static void main(String[] args) throws InterruptedException, IOException {
        String excelName = "一线业务问题答疑.xlsx";
        InputStream stream = ResourceUtil.getStream(excelName);
        ExcelReader reader = ExcelUtil.getReader(stream);
        List<List<String>> rows = CollUtil.newArrayList();
        for (int i = 0; i < reader.getSheetCount(); i++) {
            parserSheet(reader, i, rows);
        }

        //通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter("knowledge.xlsx");
        //一次性写出内容，强制输出标题
        writer.write(rows, true);
        //关闭writer，释放内存
        writer.close();
    }

    private static void parserSheet(ExcelReader reader, Integer index, List<List<String>> rows) {
        reader.setSheet(index);
        String sheetName = reader.getSheetNames().get(index);
        List<List<Object>> read = reader.read();
        List<Record> records = Lists.newArrayList();
        for (int i = 0; i < read.size(); i++) {
            if (i == 0) {
                continue;
            }
            List<Object> objects = read.get(i);
            Record record = new Record();
            if (objects == null || objects.get(0) == null) {
                continue;
            }
            record.setType(objects.get(0).toString());
            record.setOwner(objects.get(1).toString());
            record.setQuestion(objects.get(2).toString());
            record.setResolver(objects.get(3).toString());
            record.setAnswer(objects.get(4).toString());
            records.add(record);
        }
        log.info("recordListSize:{}", records.size());


        rows.add(CollUtil.newArrayList("分段标题", "分段内容", "问题"));
        for (Record record : records) {
            String title = sheetName + "-" + record.getQuestion();
            String answer = record.getAnswer();
            rows.add(CollUtil.newArrayList(title, answer));
        }
    }
}
