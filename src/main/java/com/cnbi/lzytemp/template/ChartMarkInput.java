package com.cnbi.lzytemp.template;

import com.cnbi.lzytemp.utils.SortUtils;
import com.cnbi.lzytemp.utils.XSLXUtils;
import org.apache.commons.io.IOUtils;
import org.docx4j.TraversalUtil;
import org.docx4j.dml.chart.*;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.RangeFinder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.SpreadsheetML.TablePart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorkbookPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.*;
import org.xlsx4j.sml.*;

import javax.xml.bind.JAXBElement;
import java.io.*;
import java.util.*;

/**
 * @Title: CharMarkUtil：书签以及图表插入
 * @Description:
 * @author: lzy
 * @date: 2019/8/15 16:58
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
     * @Author: lzy
     * @Date: 2019/8/22
     */
    public void insertChart(WordprocessingMLPackage wordMLPackage, Map<String, Object> data) throws Exception {
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        RelationshipsPart rp = documentPart.getRelationshipsPart();
        List<Relationship> relationships = rp.getRelationshipsByType(Namespaces.SPREADSHEETML_CHART);
        List<String> ids = new ArrayList<>();
        for (Relationship relationship : relationships) {
            String id = relationship.getId();
            ids.add(id);
        }
        //将获取的ID进行排序
        ids = SortUtils.listSort(ids);
        int charNum = 0;
        for (String id : ids) {
            //通过ID获取
            Relationship relationship = rp.getRelationshipByID(id);
            Chart chart = (Chart) rp.getPart(relationship);
            CTChartSpace chartSpace = chart.getContents();
            /*修改Excel文件*/
            String edataid = chartSpace.getExternalData().getId();
            EmbeddedPackagePart epp = (EmbeddedPackagePart) chart.getRelationshipsPart().getPart(edataid);
            ByteArrayInputStream bais = new ByteArrayInputStream(epp.getBytes());
            SpreadsheetMLPackage epack = XSLXUtils.loadExcelPackage(bais, null);
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
     * @Author: lzy
     * @Date: 2019/8/22
     */
    public int xlsheet(SpreadsheetMLPackage sheetMLP, Map<String, Object> data, int charNum) throws Docx4JException {
        List<List<String[]>> listData = (List<List<String[]>>) data.get("chart");
        List<String[]> datas = listData.get(charNum);
        WorkbookPart workbook = sheetMLP.getWorkbookPart();
        Sheets sheets = workbook.getContents().getSheets();
        Sheet sheet = sheets.getSheet().get(0);
        WorksheetPart wsPart = (WorksheetPart) workbook.getRelationshipsPart().getPart(sheet.getId());
        Worksheet worksheet = wsPart.getContents();
        /*如果sheet1.xml中包含tablePart，则需要找到对应的table1.xml修改其中的列名称*/
        CTTableParts tableparts = worksheet.getTableParts();
        if (tableparts != null) {
            CTTablePart tablepart = tableparts.getTablePart().get(0);
            TablePart tblpart = (TablePart) wsPart.getRelationshipsPart().getPart(tablepart.getId());
            List<CTTableColumn> cols = tblpart.getContents().getTableColumns().getTableColumn();
            for (int i = 0; i < cols.size(); i++) {
                //这个列数是选取区域的列数
                CTTableColumn col = cols.get(i);
                col.setName(datas.get(0)[i]);
                String colName = col.getName();
                if (colName == "" || colName == null) {
                    col.setName(" ");
                } else {
                    col.setName(colName);
                }
            }
        }
        /*修改sheet1.xml中的单元格，并替换共享字符串中对应的值*/
        List<CTRst> siList = workbook.getSharedStrings().getContents().getSi();
        List<Row> rows = worksheet.getSheetData().getRow();
        //盛放已经存在的字符串下标，防止重复
        List<Integer> indexs = new ArrayList<Integer>();
        for (int i = 0; i < datas.size(); i++) {
            List<Cell> cells = rows.get(i + 1).getC();
            for (int j = 0; j < cells.size(); j++) {
                Cell oldCell = cells.get(j);
                String newStr = datas.get(i)[j];
                STCellType type = oldCell.getT();
                //共享字符串
                if (type.equals(STCellType.S)) {
                    //对应共享字符串中的下标
                    int index = Integer.valueOf(oldCell.getV());
                    //已有该字符串，直接创建字符串单元格，否则替换共享字符串
                    if (indexs.contains(index)) {
                        Cell newCell = XSLXUtils.createCell(i, j, newStr);
                        cells.set(j, newCell);
                    } else {
                        indexs.add(index);
                        siList.get(index).getT().setValue(newStr);
                    }
                } else {
                    Cell newCell = XSLXUtils.createCell(i + 1, j, newStr);
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
     * @Author: lzy
     * @Date: 2019/8/22
     */
    public void updateChart(CTChartSpace chartSpace, Map<String, Object> data, int charNum) {
        List<List<String[]>> listData = (List<List<String[]>>) data.get("chart");
        List<String[]> datas = listData.get(charNum);
        List<Object> list = chartSpace.getChart().getPlotArea().getAreaChartOrArea3DChartOrLineChart();
        for (int i = 0, size = list.size(); i < size; i++) {
            Object obj = list.get(i);
            List<SerContent> serlist = ((ListSer) obj).getSer();
            if (obj instanceof org.docx4j.dml.chart.CTScatterChart) {
                //散点图
                for (int j = 0; j < serlist.size(); j++) {
                    SerContentXY ser = (SerContentXY) serlist.get(j);
                    CTStrVal strv = ser.getTx().getStrRef().getStrCache().getPt().get(0);
                    //strv.setV(datas.get(0)[j+1]);
                    strv.setV(strv.getV());
                    //类别名称;
                    List<CTNumVal> numlist = ser.getXVal().getNumRef().getNumCache().getPt();
                    //List<CTStrVal> strList = ser.getXVal().getStrRef().getStrCache().getPt();
                    //类别的值
                    List<CTNumVal> vallist = ser.getYVal().getNumRef().getNumCache().getPt();
                    for (int k = 0; k < numlist.size(); k++) {
                        vallist.get(k).setV(datas.get(k)[list.size() == 1 ? j + 1 : i + 1]);
                        numlist.get(k).setV(datas.get(k)[0]);
                    }
                }
            } else {
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
     * @Author: lzy
     * @Date: 2019/8/22
     */
    public void insertWords(Map<String, Object> data, WordprocessingMLPackage wPackage) throws Exception {
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
                if (bm.getName().startsWith("img")) {
                    InputStream is = new FileInputStream(new File((String.valueOf(content.get(bm.getName())))));
                    byte[] bytes = IOUtils.toByteArray(is);
                    BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wPackage, bytes);
                    int id = (int) (Math.random() * 10000);
                    // filenameHint 文件名称  altText 图片说明 id1 图片在文档中的在文档中的唯一标识 id2 图片在文档中的在文档中其他的唯一标识
                    // 最有一个是限制图片的宽度，缩放的依据
                    Inline inline = imagePart.createImageInline(null, null, id, id * 2, false, 3200);
                    //Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, false);
                    // 获取该书签的父级段落
                    P p = (P) (bm.getParent());
                    ObjectFactory factory = new ObjectFactory();
                    // R对象是匿名的复杂类型，然而我并不知道具体啥意思，估计这个要好好去看看ooxml才知道
                    R run = factory.createR();
                    // drawing理解为画布
                    Drawing drawing = factory.createDrawing();
                    drawing.getAnchorOrInline().add(inline);
                    run.getContent().add(drawing);
                    p.getContent().add(run);
                }
                ObjectFactory factory = new ObjectFactory();
                // 添加到了标签处
                P p = (P) (bm.getParent());
                //rsidPR = p.getRsidR();
                R r = factory.createR();
                Map map = traversingMark(p, content);
                List list = (List) map.get("list");
                if (list.size() > 1) {
                    content = deleteMb(list, content);
                }
            }
        }
    }

    /**
     * @Description: saveWordPackage：保存文档
     * @param: wordPackage
     * @param: filePath
     * @Return void
     * @Author: lzy
     * @Date: 2019/8/19
     */
    public void saveWordPackage(WordprocessingMLPackage wordPackage, String filePath) throws Exception {
        File file = new File(filePath);
        File fileParent = file.getParentFile();
        if (!fileParent.exists()) {
            fileParent.mkdirs();
        }
        wordPackage.save(new File(filePath));
    }

    /**
     * @Description: traversingMark:对在一行（P）中的书签进行遍历替换数据
     * @param: p一行所在的P
     * @param: content：源数据
     * @Return java.util.Map
     * @Author: lzy
     * @Date: 2019/9/11
     */
    public Map traversingMark(P p, Map content) {
        Map<String, Object> map = new HashMap<>(2);
        // 提取书签并创建书签的游标
        RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
        new TraversalUtil(p, rt);
        List<CTBookmark> starts = rt.getStarts();
        int rNum = 0;
        Map<Integer, R> mapR = new HashMap<>();
        List<Object> content2 = p.getContent();
        for (Object o : content2) {
            if (o instanceof R) {
                mapR.put(rNum, (R) o);
                rNum++;
            }
        }
        for (int i = 0; i < rNum; i++) {
            Text t = new Text();
            t.setValue((String) content.get(starts.get(i).getName()));
            R r1 = mapR.get(i);
            List<Object> content3 = r1.getContent();
            for (Object o1 : content3) {
                if (o1 instanceof JAXBElement) {
                    ((JAXBElement) o1).setValue(t.getValue());
                }
            }
        }
        map.put("p", p);
        map.put("list", starts);
        return map;
    }

    /**
     * @Description: deleteMb:删除已经替换的书签
     * @param: starts：已经替换过的书签
     * @param: content：源数据
     * @Return java.util.Map
     * @Author: lzy
     * @Date: 2019/9/11
     */
    public Map deleteMb(List<CTBookmark> starts, Map content) {
        Iterator<String> iter = content.keySet().iterator();
        while (iter.hasNext()) {
            String key = iter.next();
            for (CTBookmark start : starts) {
                if (key.equals(start.getName())) {
                    iter.remove();
                }
            }
        }
        return content;
    }
}