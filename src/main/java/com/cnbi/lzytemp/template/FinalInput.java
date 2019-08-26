package com.cnbi.lzytemp.template;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Description:模板导出
 * @Author: cnbilzy
 * @Date: 2019/8/22
 */
public class FinalInput {
    public static WordprocessingMLPackage output(String inputfilepath, Map<String, Object> data) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath));
        //向表格中添加数据
        CharMarkInput charMarkInput = new CharMarkInput();
        charMarkInput.insertChart(wordMLPackage, data);
        //替换书签信息
        charMarkInput.insertWords(data,wordMLPackage);
        TabTempInput tabTempInput = new TabTempInput();
        tabTempInput.addTabData(wordMLPackage,data);
        return wordMLPackage;
    }

    public static void main(String[] args) throws Exception {
        URL url =Thread.currentThread().getContextClassLoader().getResource("t.docx");
        String inputfilepath = url.toString().substring(5);
        SimpleDateFormat sdf = new SimpleDateFormat("", Locale.SIMPLIFIED_CHINESE);
        sdf.applyPattern("yyyy年MM月dd日HH时mm分ss秒");
        String outputfilepath = "F:/docx4j/test_" + sdf.format(System.currentTimeMillis()) + ".docx";
        WordprocessingMLPackage output = output(inputfilepath, data());
        CharMarkInput charMarkInput = new CharMarkInput();
        charMarkInput.saveWordPackage(output,outputfilepath);
    }
    public static Map<String,Object> data(){
        //添加的数据
        Map<String, Object> data = new HashMap<>(3);
        //表格数据
        List<List<List>> tableLists =  new ArrayList<>();
        List<List> tableList1 =  new ArrayList<>();
        List<List> tableList2 =  new ArrayList<>();
        List<List> rowList1 =  new ArrayList<>();
        List<List> rowList2 =  new ArrayList<>();
        List<List> rowList3 =  new ArrayList<>();
        List<List> rowList4 =  new ArrayList<>();
        List<String> list1 = new ArrayList<String>();
        List<String> list2 = new ArrayList<String>();
        list1.add("11");
        list1.add("12");
        list1.add("13");
        list1.add("14");
        list2.add("21");
        list2.add("22");
        list2.add("23");
        list2.add("24");
        list2.add("25");
        list2.add("26");
        List<String> list3 = new ArrayList<String>();
        List<String> list4 = new ArrayList<String>();
        list3.add("31");
        list3.add("32");
        list3.add("33");
        list3.add("34");
        list4.add("41");
        list4.add("42");
        list4.add("43");
        list4.add("44");
        list4.add("45");
        list4.add("46");
        rowList1.add(list1);
        rowList2.add(list2);
        rowList3.add(list3);
        rowList4.add(list4);
        tableList1.add(rowList1);
        tableList1.add(rowList2);
        tableList2.add(rowList3);
        tableList2.add(rowList4);
        tableLists.add(tableList1);
        tableLists.add(tableList2);
        data.put("table", tableLists);
        //书签数据
        Map<String,String> content = new HashMap<>();
        content.put("table1","表格1");
        content.put("table2","表格2");
        data.put("bookmark",content);
        //图表数据
        List<List<String[]>> listChart = new ArrayList<>();

        List<String[]> charList1 = new ArrayList<>();
        charList1.add(new String[]{"甲1", "2", "1", "2"});
        charList1.add(new String[]{"乙1", "3", "2", "3"});
        charList1.add(new String[]{"丙1", "4", "3", "4"});
        charList1.add(new String[]{"丁1", "4", "3", "4"});
        List<String[]> charList2 = new ArrayList<>();
        charList2.add(new String[]{"甲2", "2", "1", "2"});
        charList2.add(new String[]{"乙2", "3", "2", "3"});
        charList2.add(new String[]{"丙2", "4", "3", "4"});
        charList2.add(new String[]{"丁2", "4", "3", "4"});
        listChart.add(charList2);
        listChart.add(charList1);
        data.put("chart",listChart);
        return data;

    }
}
