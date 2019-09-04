package com.cnbi.lzytemp.template;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @Title: TabTempInput
 * @Description:
 * @author: cnbilzy
 * @date: 2019/8/20 15:04
 * @Copyright: 2019 www.cnbisoft.com Inc. All rights reserved
 * 水平合并标签：gridSpan  属性：w:val等于合并项
 * 垂直合并标签：vMerge 属性：w:val等于restart表示垂直合并的开始
 */
public class TabTempInput {

    /**
     * @Description: addTabData 向模板的表格中添加数据
     * @param: wordMLPackage
     * @param: data
     * K：table
     * V: List<List<List<List>>> tableLists  ：tableLists.get(a).get(b).get(c).get(c)
     * @Return void
     * @Author: cnbilzy
     * @Date: 2019/8/22
     */
    public void addTabData(WordprocessingMLPackage wordMLPackage,Map<String, Object> data) throws Docx4JException {

        if (data.containsKey("table")) {
            //获取所有的表格
            List<Tbl> table = getTable(wordMLPackage);
            int tbNum = 0;
            //获取单个表
            for (Tbl tbl: table) {
                int rowNum = 0;
                //获取所有的行
                List<Tr> tblAllTr = getTblAllTr(tbl);
                for (Tr tr : tblAllTr) {
                    int startNum = 0;
                    //获取所有的单元格
                    List<Tc> trAllCell = getTrAllCell(tr);
                    for (Tc tc : trAllCell) {
                        startNum = setTcContent(tc, data, rowNum, startNum, tbNum);
                    }
                    rowNum++;
                }
                tbNum++;
            }
        }
    }

    /**
     * @Description: getTable:获取所有的表格
     * @param: wordMLPackage
     * @Return java.util.List<org.docx4j.wml.Tbl>
     * @Author: cnbilzy
     * @Date: 2019/8/20
     */
    public List<Tbl> getTable(WordprocessingMLPackage wordMLPackage) throws Docx4JException {
        MainDocumentPart mainDocPart = wordMLPackage.getMainDocumentPart();
        List<Object> objList = getAllElementFromObject(mainDocPart, Tbl.class);
        if (objList == null) {
            return null;
        }
        List<Tbl> tblList = new ArrayList<Tbl>();
        for (Object obj : objList) {
            if (obj instanceof Tbl) {
                Tbl tbl = (Tbl) obj;
                tblList.add(tbl);
            }
        }
        return tblList;
    }
    /**
     * @Description: getAllElementFromObject:获取指定元素
     * @param: obj
     * @param: toSearch
     * @Return java.util.List<java.lang.Object>
     * @Author: cnbilzy
     * @Date: 2019/8/20
     */
    public List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement) {
            obj = ((JAXBElement<?>) obj).getValue();
        }
        if (obj.getClass().equals(toSearch)) {
            result.add(obj);
        }
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }
    /**
     * @Description: getTblAllTr :获取所有的行
     * @param: tbl
     * @Return java.util.List<org.docx4j.wml.Tr>
     * @Author: cnbilzy
     * @Date: 2019/8/20
     */
    public List<Tr> getTblAllTr(Tbl tbl) {
        List<Object> objList = getAllElementFromObject(tbl, Tr.class);
        List<Tr> trList = new ArrayList<Tr>();
        if (objList == null) {
            return trList;
        }
        for (Object obj : objList) {
            if (obj instanceof Tr) {
                Tr tr = (Tr) obj;
                trList.add(tr);
            }
        }
        return trList;
    }
    /**
     * @Description: getTrAllCell:获取所有的单元格
     * @param: tr
     * @Return java.util.List<org.docx4j.wml.Tc>
     * @Author: cnbilzy
     * @Date: 2019/8/20
     */
    public List<Tc> getTrAllCell(Tr tr) {
        List<Object> objList = getAllElementFromObject(tr, Tc.class);
        List<Tc> tcList = new ArrayList<Tc>();
        if (objList == null) {
            return tcList;
        }
        for (Object tcObj : objList) {
            if (tcObj instanceof Tc) {
                Tc objTc = (Tc) tcObj;
                tcList.add(objTc);
            }
        }
        return tcList;
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
    /**
     * @Description: setTcContent:设置单元格内容
     * @param: tc
     * @param: rowNum :行数
     * @param: startNum :当前单元格在本行所有不跨行、列的单元格的位置
     * @param: conten     Map<String, List> content 插入的数据
     * @Return void
     * @Author: cnbilzy
     * @Date: 2019/8/20
     */
    public int setTcContent(Tc tc, Map<String, Object> content , int rowNum , int startNum, int tbNum) {
        //一行数据为List
        //List<List<List<List>>> tableList = (List<List<List<List>>>) content.get("table");
        List<List<List<String>>> tableList = (List<List<List<String>>>) content.get("table");

        List<Object> pList = tc.getContent();
        //一行数据为List
        //List<List> conList = null;
        List<String> conList = null;
        if (rowNum > 0) {
            conList =  tableList.get(tbNum).get(rowNum-1);
        }
        
        P p = null;
        if (pList != null && pList.size() > 0) {
            if (pList.get(0) instanceof P) {
                p = (P) pList.get(0);
            }
        } else {
            p = new P();
            tc.getContent().add(p);
        }
        //获取段落的内容
        //String tContent=this.getElementContent(p);
        R run = new R();
        p.getContent().add(run);
        if (content != null) {

            //设置单元格中的值 是否存在水平、垂直合并
            TcPrInner.VMerge vm = tc.getTcPr().getVMerge();
            TcPrInner.GridSpan gridSpan = tc.getTcPr().getGridSpan();
            if ( vm == null && gridSpan == null && p.getContent() != null) {
                //清除获取单元格的内容 从第二行开始
                if (rowNum > 0 && p.getContent() != null) {
                    //p.getContent().clear();
                    for (Object o : pList) {
                        if (o instanceof P){
                            List<Object> content1 = ((P) o).getContent();
                            for (Object o1 : content1) {
                                if (o1 instanceof R){
                                    List<Object> content2 = ((R) o1).getContent();
                                    for (Object o2 : content2) {
                                        if (o2 instanceof  JAXBElement){
                                            //((JAXBElement) o2).setValue(conList.get(0).get(startNum).toString()); 一行数据一个List
                                            ((JAXBElement) o2).setValue(conList.get(startNum));
                                            startNum++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
        }

        return startNum;
    }

    
}
