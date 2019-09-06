package com.cnbi.lzytemp.template;

import com.cnbi.lzytemp.utils.SortUtils;
import com.cnbi.lzytemp.utils.XSLXUtils;
import org.docx4j.TraversalUtil;
import org.docx4j.dml.chart.*;
import org.docx4j.finders.RangeFinder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.SpreadsheetML.TablePart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorkbookPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.*;
import org.xlsx4j.sml.*;

import javax.xml.bind.JAXBElement;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @Title: CharMarkUtil：书签以及图表插入
 * @Description:
 * @author: cnbilzy
 * @date: 2019/8/15 16:58
 * @Copyright: 2019 www.cnbisoft.com Inc. All rights reserved
 * 注意：本内容仅限于安徽经邦软件有限公司内部传阅，禁止外泄以及用于其他的商业目的
 */
public class ChartMarkInput {
    /**
     * @Description: insertChart :向图表中添加数据
     * @param: wordMLPackage
     * @param: data
     * 图表数据类型为：Map<String,Object> data：
     * K为chart
     * V为List<List<String[]>> listData：listData.get(a).get(b)[c] a表示地几个图表，b是第几行
     * @Return void
     * @Author: cnbilzy
     * @Date: 2019/8/22
     */
    public void insertChart(WordprocessingMLPackage wordMLPackage ,Map<String,Object> data) throws Exception {
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
            RelationshipsPart rp = documentPart.getRelationshipsPart();
            List<Relationship> relationships = rp.getRelationshipsByType(Namespaces.SPREADSHEETML_CHART);
            List<String> ids = new ArrayList<>();
            for (Relationship relationship : relationships) {
                String id = relationship.getId();
                ids.add(id);
            }
            //获取ID进行排序
            ids = SortUtils.listSort(ids);
            int charNum = 0;
            /*for (int i = ids.size()-1; i >= 0 ; i--) {
                String id = ids.get(i);
                //通过ID获取
                Relationship relationship = rp.getRelationshipByID(id);
                Chart chart = (Chart) rp.getPart(relationship);
                CTChartSpace chartSpace = chart.getContents();
                *//*修改Excel文件*//*
                String edataid = chartSpace.getExternalData().getId();
                EmbeddedPackagePart epp = (EmbeddedPackagePart)chart.getRelationshipsPart().getPart(edataid);
                ByteArrayInputStream bais = new ByteArrayInputStream(epp.getBytes());
                SpreadsheetMLPackage epack =  XSLXUtils.loadExcelPackage(bais, null);
                int num = xlsheet(epack, data, charNum);
                File filetemp = File.createTempFile("temp", "xlsx");
                epack.save(new FileOutputStream(filetemp), 1, null);
                epp.setBinaryData(new FileInputStream(filetemp));
                //修改xml数据
                updateChart(chartSpace, data, charNum);
                charNum = num;
            }*/
            for (String id : ids) {
                //通过ID获取
                Relationship relationship = rp.getRelationshipByID(id);
                Chart chart = (Chart) rp.getPart(relationship);
                CTChartSpace chartSpace = chart.getContents();
                /*修改Excel文件*/
                String edataid = chartSpace.getExternalData().getId();
                EmbeddedPackagePart epp = (EmbeddedPackagePart)chart.getRelationshipsPart().getPart(edataid);
                ByteArrayInputStream bais = new ByteArrayInputStream(epp.getBytes());
                SpreadsheetMLPackage epack =  XSLXUtils.loadExcelPackage(bais, null);
                int num = xlsheet(epack, data, charNum);
                File filetemp = File.createTempFile("temp", "xlsx");
                epack.save(new FileOutputStream(filetemp), 1, null);
                epp.setBinaryData(new FileInputStream(filetemp));
                //修改xml数据
                updateChart(chartSpace, data, charNum);
                charNum = num;
            }

    }
    /**
     * @Description: xlsheet：替换chart对应的表格数据
     * @param: sheetMLP
     * @param: data
     * @param: charNum
     * @Return int
     * @Author: cnbilzy
     * @Date: 2019/8/22
     */
    public int xlsheet(SpreadsheetMLPackage sheetMLP, Map<String, Object> data, int charNum) throws Docx4JException  {
        List<List<String[]>> listData = (List<List<String[]>>) data.get("chart");
        List<String[]> datas = listData.get(charNum);
        WorkbookPart workbook = sheetMLP.getWorkbookPart();
        Sheets sheets = workbook.getContents().getSheets();
        Sheet sheet = sheets.getSheet().get(0);
        WorksheetPart wsPart = (WorksheetPart)workbook.getRelationshipsPart().getPart(sheet.getId());
        Worksheet worksheet = wsPart.getContents();
        /*如果sheet1.xml中包含tablePart，则需要找到对应的table1.xml修改其中的列名称*/
        CTTableParts tableparts = worksheet.getTableParts();
        if(tableparts != null){
            CTTablePart tablepart = tableparts.getTablePart().get(0);
            TablePart tblpart = (TablePart)wsPart.getRelationshipsPart().getPart(tablepart.getId());
            List<CTTableColumn> cols = tblpart.getContents().getTableColumns().getTableColumn();
            for(int i = 0;i < cols.size();i++){

                //这个列数是选取区域的列数
                CTTableColumn col = cols.get(i);
                col.setName(datas.get(0)[i]);
                String colName = col.getName();
                if(colName == "" || colName == null){
                    col.setName(" ");
                }else{
                    col.setName(colName);
                }
            }
        }
        /*修改sheet1.xml中的单元格，并替换共享字符串中对应的值*/
        List<CTRst> siList = workbook.getSharedStrings().getContents().getSi();
        List<Row> rows = worksheet.getSheetData().getRow();
        //盛放已经存在的字符串下标，防止重复
        List<Integer> indexs = new ArrayList<Integer>();
        for(int i = 0;i < datas.size();i++){
            List<Cell> cells = rows.get(i+1).getC();
            for(int j = 0; j < cells.size();j++){
                Cell oldCell = cells.get(j);
                String newStr = datas.get(i)[j];
                //newStr = formatNum(newStr);
                STCellType type = oldCell.getT();
                //共享字符串
                if(type.equals(STCellType.S)){
                    //对应共享字符串中的下标
                    int index = Integer.valueOf(oldCell.getV());
                    //已有该字符串，直接创建字符串单元格，否则替换共享字符串
                    if(indexs.contains(index)){
                        Cell newCell = XSLXUtils.createCell(i, j, newStr);
                        cells.set(j, newCell);
                    }else{
                        indexs.add(index);
                        siList.get(index).getT().setValue(newStr);
                    }
                }else{
                    Cell newCell = XSLXUtils.createCell(i+1, j, newStr);
                    cells.set(j, newCell);
                }
            }
        }
        charNum++;
        return charNum;
    }
    /**
     * @Description: updateChart:更新图表数据
     * @param: chartSpace
     * @param: data
     * @param: charNum
     * @Return void
     * @Author: cnbilzy
     * @Date: 2019/8/22
     */
    public void updateChart(CTChartSpace chartSpace,  Map<String, Object> data, int charNum){
        List<List<String[]>> listData = (List<List<String[]>>) data.get("chart");
        List<String[]> datas = listData.get(charNum);
        List<Object> list = chartSpace.getChart().getPlotArea().getAreaChartOrArea3DChartOrLineChart();
        for(int i = 0, size = list.size(); i < size ; i++){
            Object obj = list.get(i);
            List<SerContent> serlist = ((ListSer)obj).getSer();
            if(obj instanceof org.docx4j.dml.chart.CTScatterChart){
                //散点图
                for(int j = 0; j < serlist.size(); j++){
                    SerContentXY ser = (SerContentXY) serlist.get(j);
                    CTStrVal strv =  ser.getTx().getStrRef().getStrCache().getPt().get(0);
                    //strv.setV(datas.get(0)[j+1]);
                    strv.setV(strv.getV());
                    //类别名称;
                    List<CTNumVal> strlist = ser.getXVal().getNumRef().getNumCache().getPt();
                    //类别的值
                    List<CTNumVal> vallist = ser.getYVal().getNumRef().getNumCache().getPt();
                    for(int k = 0; k < strlist.size(); k++ ){
                        vallist.get(k).setV(datas.get(k)[list.size() == 1 ? j + 1 : i + 1]);
                        strlist.get(k).setV(datas.get(k)[0]);
                    }
                }
            }else {
                //柱形图、条形图、折线图、饼图、面积图、圆环图、雷达图等
                for (int j = 0; j < serlist.size(); j++) {
                    SerContent ser = serlist.get(j);
                    CTStrVal strv = ser.getTx().getStrRef().getStrCache().getPt().get(0);
                    strv.setV(strv.getV());
                    //类别名称
                    List<CTStrVal> strlist = ser.getCat().getStrRef().getStrCache().getPt();
                    //类别的值
                    List<CTNumVal> vallist = ser.getVal().getNumRef().getNumCache().getPt();
                    for (int k = 0; k < strlist.size(); k++) {
                        vallist.get(k).setV(datas.get(k)[list.size() == 1 ? j + 1 : i + 1]);
                        strlist.get(k).setV(datas.get(k)[0]);
                        //vallist.get(k).setV(datas.get(k)[j+1]);
                    }
                }
            }
        }
    }

    /**
     * @Description: insertWords:向书签处插入数据
     * @param: data Map<String,Map> data
     * K： bookmark
     * V ：Map :K:书签名 V：插入的内容
     * @param: wPackage
     * @Return void
     * @Author: cnbilzy
     * @Date: 2019/8/22
     */
    public void insertWords(Map<String,Object> data, WordprocessingMLPackage wPackage){
            Map content = (Map) data.get("bookmark");
            // 提取正文
            MainDocumentPart mainDocumentPart = wPackage.getMainDocumentPart();
            Document wmlDoc = (Document) mainDocumentPart.getJaxbElement();
            Body body = wmlDoc.getBody();
            // 提取正文中所有段落
            List<Object> paragraphs = body.getContent();
            // 提取书签并创建书签的游标
            RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
            new TraversalUtil(paragraphs, rt);
            // 遍历书签
            for (CTBookmark bm : rt.getStarts()) {
                // 这儿可以对单个书签进行操作，也可以用一个map对所有的书签进行处理
                if (content.containsKey(bm.getName())) {
                    ObjectFactory factory = new ObjectFactory();
                    // 添加到了标签处
                    P p = (P) (bm.getParent());
                    R r = factory.createR();
                    List<Object> content1 = p.getContent();
                    Text t = new Text();
                    t.setValue((String) content.get(bm.getName()));
                    List<Object> content2 = p.getContent();
                    for (Object o : content2) {
                        if (o instanceof R){
                            List<Object> content3 = ((R) o).getContent();
                            for (Object o1 : content3) {
                                if (o1 instanceof JAXBElement){
                                    ((JAXBElement) o1).setValue(t);
                                }
                            }

                        }
                    }
                    p.getContent().add(r);
                }
            }

    }
    /**
     * @Description: saveWordPackage：保存文档
     * @param: wordPackage
     * @param: filePath
     * @Return void
     * @Author: cnbilzy
     * @Date: 2019/8/19
     */
    public void saveWordPackage(WordprocessingMLPackage wordPackage, String filePath) throws Exception {
        File file = new File(filePath);
        File fileParent = file.getParentFile();
        if(!fileParent.exists()){
            fileParent.mkdirs();
        }
        wordPackage.save(new File(filePath));
    }






}
