package com.cnbi.lzytemp.template;
import com.cnbi.lzytemp.utils.DOCUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import java.text.SimpleDateFormat;
import java.util.*;
import static org.docx4j.wml.JcEnumeration.RIGHT;
/**
 * @Description:模板导出
 * @Author: lzy
 * @Date: 2019/8/22
 */
public class FinalInput {
    final static String CHART = "chart";
    final static String BOOKMARK = "bookmark";
    final static String TABLE = "table";
    public WordprocessingMLPackage output(String inputfilepath, Map<String, Object> data) throws Exception {
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
            tabTempInput.addTabData(wordMLPackage, data);
        }

        return wordMLPackage;
    }
    public static void main(String[] args) throws Exception {
        String property = System.getProperty("user.dir");
        //String inputfilepath = property+"\\src\\main\\resources\\t.docx";
        String inputfilepath = "C:\\Users\\Administrator\\Desktop\\t.docx";
        SimpleDateFormat sdf = new SimpleDateFormat("", Locale.SIMPLIFIED_CHINESE);
        sdf.applyPattern("yyyy年MM月dd日HH时mm分ss秒");
        String outputfilepath = "F:/docx4j/test_" + sdf.format(System.currentTimeMillis()) + ".docx";
        /*WordprocessingMLPackage output = output(inputfilepath, data());
        //添加页眉
        DOCUtils.createHeader(output, "(1001) 天津食品集团有限公司（合并）2019年8月风险报告", RIGHT);
        ChartMarkInput charMarkInput = new ChartMarkInput();
        charMarkInput.saveWordPackage(output, outputfilepath);*/
    }
    public static Map<String, Object> data() {
        //添加的数据
        Map<String, Object> data = new HashMap<>(3);
        //表格数据
        TreeMap<String, List<List>> tableTree1 = new TreeMap<>();
        TreeMap<String, List<List>> tableTree2 = new TreeMap<>();
        List<TreeMap<String, List<List>>> tableLists = new ArrayList<>();
        List<List> tableList1 = new ArrayList<>();
        List<List> tableList2 = new ArrayList<>();
        List<String> list1 = new ArrayList<>();
        List<String> list2 = new ArrayList<>();
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
        List<String> list3 = new ArrayList<>();
        List<String> list4 = new ArrayList<>();
        list3.add("31");
        list3.add("32");
        list3.add("33");
        list3.add("34");
        list3.add(null);
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
        tableTree1.put("1", tableList1);
        tableTree2.put("2", tableList2);
        tableLists.add(tableTree1);
        tableLists.add(tableTree2);
        data.put("table", tableLists);
        //书签数据
        Map<String, String> content = new HashMap<>();
        content.put("table1", "表格1");
        content.put("table2", "表格2");
        content.put("table3", "表格3");
        content.put("mulu1", "目录1");
        content.put("mulu2", "目录2");
        content.put("mulu3", "目录3");
        content.put("header", "(1001) 天津食品集团有限公司（合并）2019年8月风险报告");
        content.put("text1", "Apache POI包括一系列的API，它们可以操作基于MicroSoft OLE 2 Compound Document Format的各种格式文件，" +
                "可以通过这些API在Java中读写Excel、Word等文件。他的excel处理很强大，对于word还局限于读取，目前只能实现一些简单文件的" +
                "操作，不能设置样式Apache POI包括一系列的API，它们可以操作基于MicroSoft OLE 2 Compound Document Format的各种格式" +
                "文件，可以通过这些API在Java中读写Excel、Word等文件。他的excel处理很强大，对于word还局限于读取，目前只能实现一些简" +
                "单文件的操作，不能设置样式");
        content.put("text2", "用XML做就很简单了。Word从2003开始支持XML格式，大致的思路是先用office2003或者2007编辑好word的样式，" +
                "然后另存为xml，将xml翻译为FreeMarker模板，最后用java来解析FreeMarker模板并输出Doc。经测试这样方式生成的word文" +
                "档完全符合office标准，样式、内容控制非常便利，打印也不会变形，生成的文档和office中编辑文档完全一样。用XML做就很简单" +
                "了。Word从2003开始支持XML格式，大致的思路是先用office2003或者2007编辑好word的样式，然后另存为xml，将xml翻译为" +
                "FreeMarker模板，最后用java来解析FreeMarker模板并输出Doc。经测试这样方式生成的word文档完全符合office标准，样式、" +
                "内容控制非常便利，打印也不会变形，生成的文档和office中编辑文档完全一样。");
        content.put("img1", "D:\\Documents\\Tencent Files\\312593790\\FileRecv\\201903171.jpg");
        content.put("com", "北京首创股份有限公司");
        content.put("rpt", "2019年1-2季度经营分析报告");
        content.put("period", "2019年6月编制");
        content.put("per", "2018年12月编制");
        data.put("bookmark", content);
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
        charList3.add(new String[]{"1", "1", "2", "3"});
        charList3.add(new String[]{"2", "1", "2", "3"});
        charList3.add(new String[]{"3", "1", "2", "3"});
        charList3.add(new String[]{"4", "1", "2", "3"});
        listChart.add(charList1);
        listChart.add(charList2);
        listChart.add(charList3);
        data.put("chart", listChart);
        return data;
    }
}