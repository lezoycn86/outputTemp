package com.cnbi.lzytemp.template;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Description:模板导出
 * @Author: cnbilzy
 * @Date: 2019/8/22
 */
public class FinalInput {
    final static String CHART = "chart";
    final static String BOOKMARK = "bookmark";
    final static String TABLE = "table";
    public static WordprocessingMLPackage output(String inputfilepath, Map<String, Object> data) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath));
        //向图表中添加数据
        ChartMarkInput chartMarkInput = new ChartMarkInput();
        if (data.containsKey(CHART)) {
            chartMarkInput.insertChart(wordMLPackage, data);
        }
        //替换书签信息
        if (data.containsKey(BOOKMARK)) {
            chartMarkInput.insertWords(data, wordMLPackage);
        }
        //向表格中添加数据
        if (data.containsKey(TABLE)) {
            TabTempInput tabTempInput = new TabTempInput();
            tabTempInput.addTabData(wordMLPackage,data);
        }
        return wordMLPackage;
    }

    public static void main(String[] args) throws Exception{
        String property = System.getProperty("user.dir");
        String inputfilepath = property+"\\src\\main\\resources\\t.docx";
        SimpleDateFormat sdf = new SimpleDateFormat("", Locale.SIMPLIFIED_CHINESE);
        sdf.applyPattern("yyyy年MM月dd日HH时mm分ss秒");
        //String inputfilepath = "D:\\Downloads\\(1001) 天津食品集团有限公司（合并）2019年8月风险报告.docx";
        String outputfilepath = "F:/docx4j/test_" + sdf.format(System.currentTimeMillis()) + ".docx";
        WordprocessingMLPackage output = output(inputfilepath, data());
        ChartMarkInput charMarkInput = new ChartMarkInput();
        charMarkInput.saveWordPackage(output,outputfilepath);
    }
    public static  Map<String,Object> data(){
        //添加的数据
        Map<String, Object> data = new HashMap<>(3);
        //表格数据
        TreeMap<String,List<List>> tableTrees =  new TreeMap<>();
        List<List> tableList1 =  new ArrayList<>();
        List<List> tableList2 =  new ArrayList<>();
        List<String> list1 = new ArrayList<String>();
        List<String> list2 = new ArrayList<String>();
        list1.add("11");
        list1.add("12");
        list1.add("13");
        list1.add("14");
        list1.add("15");
        list1.add("16");
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
        list3.add("35");
        list3.add("36");
        list4.add("41");
        list4.add("42");
        list4.add("43");
        list4.add("44");
        list4.add("45");
        list4.add("46");

        tableList1.add(list1);
        tableList1.add(list2);
        tableList2.add(list3);
        tableList2.add(list4);
        //key从第几行填入数据 0是第一行
        tableTrees.put("1",tableList1);
        tableTrees.put("2",tableList2);
        data.put("table", tableTrees);
        //书签数据
        Map<String,String> content = new HashMap<>();
        content.put("table1","表格1");
        content.put("table2","表格2");
        content.put("mulu1","目录1");
        content.put("mulu2","目录2");
        content.put("mulu3","目录3");
        content.put("year","2019");
        content.put("title","年");
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
        List<String[]> charList3 = new ArrayList<>();
        charList3.add(new String[]{"甲3", "1", "1", "1"});
        charList3.add(new String[]{"乙3", "2", "2", "2"});
        charList3.add(new String[]{"丙3", "3", "3", "3"});
        charList3.add(new String[]{"丁3", "4", "4", "4"});
        listChart.add(charList1);
        listChart.add(charList2);
        listChart.add(charList3);
        data.put("chart",listChart);
        return data;

    }
}
