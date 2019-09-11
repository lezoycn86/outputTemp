package com.cnbi.lzytemp.utils;

import com.cnbi.lzytemp.base.DocxConstains;
import com.cnbi.lzytemp.base.StyleConstains;
import org.docx4j.Docx4jProperties;
import org.docx4j.UnitsOfMeasurement;
import org.docx4j.XmlUtils;
import org.docx4j.dml.*;
import org.docx4j.dml.chart.CTRelId;
import org.docx4j.dml.wordprocessingDrawing.CTEffectExtent;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.PropertyResolver;
import org.docx4j.model.structure.PageSizePaper;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.WordprocessingML.*;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.CTSdtDocPart.DocPartGallery;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.PBdr;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.TcPrInner.GridSpan;
import org.docx4j.wml.TcPrInner.VMerge;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.namespace.QName;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;


public class DOCUtils {
	
	/**功能描述：创建文档处理包对象
	 * @return  返回值：返回文档处理包对象
	 * @Description: createWordprocessingMLPackage
	 * @param:
	 * @Return org.docx4j.openpackaging.packages.WordprocessingMLPackage
	 * @Author: cnbilzy
	 * @Date: 2019/9/9
	 */
	public static WordprocessingMLPackage createWordprocessingMLPackage() throws Exception{
		return WordprocessingMLPackage.createPackage();
	}
	
	public static WordprocessingMLPackage createWordprocessingMLPackageForA3() throws InvalidFormatException {
	    String papersize= Docx4jProperties.getProperties().getProperty("docx4j.PageSize", "A3");
	    String landscapeString = Docx4jProperties.getProperties().getProperty("docx4j.PageOrientationLandscape", "false");
	    boolean landscape= Boolean.parseBoolean(landscapeString);
	    return WordprocessingMLPackage.createPackage(
				PageSizePaper.valueOf(papersize), landscape);
	}
	/**
	 * @Description: loadWordprocessingMLPackage 功能描述：加载文档
	 * @param: filePath 文档路径
	 * @Return org.docx4j.openpackaging.packages.WordprocessingMLPackage
	 * @Author: cnbilzy
	 * @Date: 2019/9/9
	 */
	public static WordprocessingMLPackage loadWordprocessingMLPackage(String filePath)throws Exception {
	  return loadWordPackage(new File(filePath), null);
	}
	
	/**
	 * @Description: loadWordPackage 功能描述：加载文档信息
	 * @param: wordFile
	 * @param: password
	 * @Return org.docx4j.openpackaging.packages.WordprocessingMLPackage
	 * @Author: cnbilzy
	 * @Date: 2019/9/9
	 */
	public static WordprocessingMLPackage loadWordPackage(File wordFile, String password) throws Exception {
		return (WordprocessingMLPackage) BaseUtils.loadOpcPackage(wordFile, password);
	}
	
	
	/**
	 * @Description: saveWordPackage 功能描述：保存文档信息
	 * @param: wordPackage    文档处理包对象
	 * @param: file           文件
	 * @Return void
	 * @Author: cnbilzy
	 * @Date: 2019/9/9
	 */
	public static void saveWordPackage(WordprocessingMLPackage wordPackage, File file) throws Exception {
		BaseUtils.savePackage(wordPackage, file);
	}
	
	/**
	 * 功能描述：设置页边距
	 * @param wordPackage 文档处理包对象
	 * @param top    上边距
	 * @param bottom 下边距
	 * @param left   左边距
	 * @param right  右边距
	 * @author cnbi
	 */
	public static void setMarginSpace(WordprocessingMLPackage wordPackage , String top , String bottom , String left , String right ){
		ObjectFactory factory = Context.getWmlObjectFactory();
		PgMar pg = factory.createSectPrPgMar();
		pg.setTop(new BigInteger(top));
		pg.setBottom(new BigInteger(bottom));
		pg.setLeft(new BigInteger(left));
		pg.setRight(new BigInteger(right));
		wordPackage.getDocumentModel().getSections().get(0).getSectPr().setPgMar(pg);
	}
	
	/**
	 * 功能描述：设置页边距，上下边距都为1440,2.54厘米，左右边距都为1797,3.17厘米
	 * @param wordPackage 文档处理包对象
	 * @author cnbi
	 */
	public static void setMarginSpace(WordprocessingMLPackage wordPackage){
		setMarginSpace(wordPackage, "1440", "1440", "1797", "1797");
	}
	
	
	
	/**
	 * 设成横版
	 * @param wordPackage
	 */
	public static void setLandscapeMode(WordprocessingMLPackage wordPackage){
		setPageOrientation(wordPackage, STPageOrientation.LANDSCAPE);
	}
	
	/**
	 * 页面版式
	 * @param wordPackage
	 * @param orient  portrait 竖版     landscape 横版
	 */
	public static void setPageOrientation(WordprocessingMLPackage wordPackage, STPageOrientation orient){
		SectPr.PgSz pgSz = wordPackage.getDocumentModel().getSections().get(0).getSectPr().getPgSz();
		pgSz.setOrient(orient);
		Long h = pgSz.getH().longValue();
		Long w = pgSz.getW().longValue();
		if((STPageOrientation.LANDSCAPE.equals(orient) &&  h > w) || (STPageOrientation.PORTRAIT.equals(orient) &&  w > h)){
			pgSz.setH(BigInteger.valueOf(w));
			pgSz.setW(BigInteger.valueOf(h));
		}
	}

	/**
	 * 功能描述：获取文档的可用宽度
	 * @param wordPackage 文档处理包对象
	 * @return            返回值：返回值文档的可用宽度
	 * @throws Exception
	 * @author cnbi
	 */
	public static int getWritableWidth(WordprocessingMLPackage wordPackage)throws Exception{
		return wordPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
	}

	/**
	 * 功能描述：往文档对象中添加相应的内容
	 * @param wordPackage  文档处理包对象
	 * @param info         需要添加的信息
	 * @param unmarshal    是否有样式，表格对象默认不用
	 * @throws Exception
	 * @author cnbi
	 */
	private static void addObject(WordprocessingMLPackage wordPackage, Object info, boolean unmarshal) {
		if (unmarshal) {
			try{
			wordPackage.getMainDocumentPart().addObject(XmlUtils.unmarshalString(String.valueOf(info)));
			}catch(JAXBException ex){
				addObject(wordPackage, info, false);
			}
		} else {
			wordPackage.getMainDocumentPart().addObject(info);
		}
	}
	/**
	 * 功能描述：往文档中添加内容
	 * @param wordPackage  文档处理包对象
	 * @param info         文档内容
	 * @throws Exception
	 * @author cnbi
	 */
	public static void addObject(WordprocessingMLPackage wordPackage, Object info) {
		addObject(wordPackage, info, true);
	}

	/**
	 * @Description: addPageBreak 功能描述：段落添加Br 页面Break(分页符)
	 * @param: wordPackage
	 * @param: sTBrType
	 * @Return org.docx4j.wml.P
	 * @Author: cnbilzy
	 * @Date: 2019/9/9
	 */
	public static P addPageBreak(WordprocessingMLPackage wordPackage, STBrType sTBrType) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		P para = factory.createP();
		Br breakObj = new Br();
		breakObj.setType(sTBrType);
		R r = creatRun();
		r.getContent().add(breakObj);
		para.getContent().add(r);
		addObject(wordPackage, para, false);
		return para;
	}

	/**
	 * 功能描述：添加带正文样式的文本段落
	 * @param wordPackage
	 * @param text
	 * @return
	 */
	public static P addNormalParagraphOfText(WordprocessingMLPackage wordPackage, String text){
		return addStyledParagraphOfText(wordPackage, "Normal", text);
	}
	/**
	 *
	 * @param wordPackage
	 * @param text
	 * @return
	 */
	public static P addBodyParagraphOfText(WordprocessingMLPackage wordPackage, String text){
		return addStyledParagraphOfText(wordPackage, StyleConstains.CNBIBody, text);
	}

	/**
	 * 功能描述：添加带样式的文本段落
	 * @param wordPackage 文档处理包对象
	 * @param styleId  样式id 或 名称
	 * @param text  文本内容
	 * @throws Exception
	 */
	public static P addStyledParagraphOfText(WordprocessingMLPackage wordPackage, String styleId, String text) {
		P p = createStyledParagraphOfText(wordPackage, styleId, text);
		addObject(wordPackage, p, false);
		return p;
	}

	/**
	 *
	 * @param wordPackage
	 * @param titlelevel Title 主标题 , Subtitle 副标题,  1.2.3
	 * @param title
	 * @return
	 * @throws Exception
	 */
	public static P setTitle(WordprocessingMLPackage wordPackage, String titlelevel, String title){
		return setTitle(wordPackage, titlelevel, title, null, -1);
	}
	public static P setTitle(WordprocessingMLPackage wordPackage, String titlelevel, String title,String mark, int index){
		P p = null;
		if("Title".equals(titlelevel)){
			p = addStyledParagraphOfText(wordPackage, "Title", title);
		}else if("Subtitle".equals(titlelevel)){
			p = addStyledParagraphOfText(wordPackage, "Subtitle", title);
		}else{
			p = addStyledParagraphOfText(wordPackage, "Heading"+titlelevel, title);
//			if(mark != null){
//				setBookMark(p, mark, index);
//			}
		}
		return p;
	}

	/**
	 * 设置书签
	 * @param p
	 * @param mark
	 * @param index
	 */
	private static void setBookMark(P p , String mark, int index){
		ObjectFactory wmlObjectFactory = Context.getWmlObjectFactory();
		CTBookmark bookmark = wmlObjectFactory.createCTBookmark();
		JAXBElement<CTBookmark> bookmarkWrapped = wmlObjectFactory.createPBookmarkStart(bookmark);
		bookmark.setName(mark);
		bookmark.setId(BigInteger.valueOf(index));
		p.getContent().add(0, bookmarkWrapped);

      CTMarkupRange markuprange = wmlObjectFactory.createCTMarkupRange();
      markuprange.setId( BigInteger.valueOf( index) );
      JAXBElement<CTMarkupRange> markuprangeWrapped = wmlObjectFactory.createPBookmarkEnd(markuprange);
      p.getContent().add(markuprangeWrapped);
	}

	/**
	 * 功能描述：设置段落内容
	 * @param wordPackage
	 * @param list
	 */
	public static P setParaRContent(WordprocessingMLPackage wordPackage, List list) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		P p = addNormalParagraphOfText(wordPackage, "");
		//块样式
		RPr rpr = factory.createRPr();
		setParaRContent(p, rpr, list);
		return p;
	}

	/**
	 * 功能描述：设置段落内容
	 * @param p 段落对象
	 * @param runProperties
	 * @param content
	 */
	public static void setParaRContent(P p, RPr runProperties, List content) {
		R run = null;
		List<Object> rList = p.getContent();
		if (rList != null && rList.size() > 0) {
		  for(int i=0,len=rList.size();i<len;i++){
		    // 清除内容(所有的r
		    p.getContent().remove(0);
		  }
		}
		run = new R();
		p.getContent().add(run);
		if (content != null && content.size()>0) {
		Text text = getText(content.get(0)+"", "preserve");

		run.setRPr(runProperties);
		run.getContent().add(text);

		for (int i = 1, len = content.size(); i < len; i++) {
			  Br br = new Br();
			  run.getContent().add(br);// 换行
			  text = getText(content.get(i)+"", "preserve");
		     run.setRPr(runProperties);
		     run.getContent().add(text);
		   }
		 }
	}

	public static Text getText(String content, String space){
		Text text = new Text();
		if(space != null){
			text.setSpace(space);
		}
		text.setValue(content);
		return text;
	}

	/**
	 * 功能描述：创建有样式的文本段落对象
	 * @param wordPackage 文档处理包对象
	 * @param styleId  样式id 或 名称
	 * @param text  文本内容
	 * @return
	 */
	public static P createStyledParagraphOfText(WordprocessingMLPackage wordPackage, String styleId, String text) {
		P styledParagraphOfText = wordPackage.getMainDocumentPart().createStyledParagraphOfText(styleId, text);
		String pStyleVal = styledParagraphOfText.getPPr().getPStyle().getVal();
		//为报告中的段落添加段间距
		if (pStyleVal.contains("BodyText")) {
			addLineSpace(styledParagraphOfText);
		}
		return styledParagraphOfText;
	}
	/**
	 * @Description: addLineSpace:设置段间距为0.5
	 * @param: p
	 * @Author: cnbilzy
	 * @Date: 2019/7/30
	 */
	public static void addLineSpace (P p) {
		//设置段间距0.5
		ObjectFactory factory = Context.getWmlObjectFactory();
		Spacing sp = factory.createPPrBaseSpacing();
		PPr paragraphPPr = factory.createPPr();
		//sp.setAfter(BigInteger.valueOf(120));
		//sp.setBefore(BigInteger.valueOf(120));
		//数值是word中的100倍
		sp.setBeforeLines(new BigInteger(String.valueOf(50)));
		sp.setAfterLines(new BigInteger(String.valueOf(50)));
		paragraphPPr.setSpacing(sp);
		p.setPPr(paragraphPPr);

	}

	/**
	 * 获取块
	 * @return
	 */
	public static R creatRun(){
		ObjectFactory factory = Context.getWmlObjectFactory();
		R r = factory.createR();
		return r;
	}

	/**
	 * 设置段落样式
	 * @param wordPackage
	 * @param p
	 * @param styleId
	 */
	public static void setStyle(WordprocessingMLPackage wordPackage, P p, String styleId){
		if (wordPackage.getMainDocumentPart().getPropertyResolver().activateStyle(styleId)) {// Style is available
			setStyle(p, styleId);
		}
	}


	public static void setStyle(P p, String styleId){
		ObjectFactory factory = Context.getWmlObjectFactory();
		PPr  pPr = getPPr(p);
		PStyle pStyle = factory.createPPrBasePStyle();
		pPr.setPStyle(pStyle);
		pStyle.setVal(styleId);
	}

	/**
	 * 功能描述：获取块样式
	 * @param r
	 * @return
	 */
	public static RPr getRPr(R r){
		RPr rpr = r.getRPr();
		if(rpr == null){
			rpr = new RPr();
			r.setRPr(rpr);
		}
		return rpr;
	}

	/**
	* 功能描述：设置字体的样式，宋体，黑色，18号
	* @param isBlod          是否加粗
	* @return                返回值：返回字体样式对象
	* @throws Exception
	* @author cnbi
	*/
	public static RPr getRPr(Boolean isBlod) {
		return getRPr("宋体", "000000", "18", STHint.EAST_ASIA, isBlod);
	}
	/**
	 * 功能描述：设置字体的样式，黑色，18号
	 * @param fontFamily      字体
	 * @param isBlod          是否加粗
	 * @return                返回值：返回字体样式对象
	 * @throws Exception
	 * @author cnbi
	 */
	public static RPr getRPr(String fontFamily, Boolean isBlod) {
		return getRPr(fontFamily, "000000", "18", STHint.EAST_ASIA, isBlod);
	}
	/**
	 * 功能描述：设置字体的样式，黑色
	 * @param fontFamily      字体
	 * @param hpsMeasureSize  字号的大小
	 * @param isBlod          是否加粗
	 * @return                返回值：返回字体样式对象
	 * @throws Exception
	 * @author cnbi
	 */
	public static RPr getRPr(String fontFamily, String hpsMeasureSize, Boolean isBlod) {
		return getRPr(fontFamily, "000000", hpsMeasureSize, STHint.EAST_ASIA, isBlod);
	}

	public static RPr getRPrForSCGFTable(String fontFamily, Boolean isBlod, Boolean flag) {
		RPr rPr = getRPr("仿宋", "000000", "18", null, isBlod);
		rPr.getRFonts().setCs("Arial");
		if(flag != null) {
		    BooleanDefaultTrue b = rPr.getB();
		    if(b == null) {
			b = new BooleanDefaultTrue();
		    }
		    b.setVal(flag);
		    rPr.setI(b);
		    rPr.setB(b);
		}
		return rPr;
	}
	/**
	 * 功能描述：设置字体的样式
	 * @param fontFamily      字体类型
	 * @param colorVal        字体颜色
	 * @param hpsMeasureSize  字号大小
	 * @param sTHint          字体格式
	 * @param isBlod          是否加粗
	 * @return                返回值：返回字体样式对象
	 * @throws Exception
	 * @author cnbi
	 */
	public static RPr getRPr(String fontFamily, String colorVal, String hpsMeasureSize, STHint sTHint, Boolean isBlod) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		RPr rPr = factory.createRPr();
		//字体
		if(fontFamily != null){
		RFonts rf = new RFonts();
		rf.setHint(sTHint);
		rf.setEastAsia(fontFamily);
		rf.setAscii(fontFamily);
		rf.setHAnsi(fontFamily);
		rPr.setRFonts(rf);
		}
		//加粗
		BooleanDefaultTrue bdt = factory.createBooleanDefaultTrue();
		if (isBlod != null) {
			bdt.setVal(isBlod);
			rPr.setB(bdt);
			rPr.setBCs(bdt);
		}
		//字体颜色
		if(colorVal != null){
			setColorForRPR(rPr, colorVal);
		}
		//字体大小
		if(hpsMeasureSize != null){
			setSizeForRPR(rPr, hpsMeasureSize);
		}
//		//底纹
//		CTShd shd = new CTShd();
//		shd.setColor();
//		rPr.setShd(shd);

		return rPr;
	}

	/**
	 * 设置块字体大小
	 * @param hpsMeasureSize
	 * @param rPr
	 */
	public static void setSizeForRPR( RPr rPr, String hpsMeasureSize) {
		HpsMeasure sz = rPr.getSz();
		if(sz == null){
			sz = new HpsMeasure();
			rPr.setSz(sz);
			rPr.setSzCs(sz);
		}
		sz.setVal(new BigInteger(hpsMeasureSize));
	}

	/**
	 * 设置块字体颜色
	 * @param rpr
	 * @param colorVar RGB值  :"000000"
	 */
	public static void setColorForRPR(RPr rpr,String colorVar){
		Color color = rpr.getColor();
		if(color == null){
			color = new Color();
			rpr.setColor(color);
		}
		color.setVal(colorVar);
	}

	/**
	 * 功能描述：设置底纹
	 * @param rpr
	 * @param shdtype
	 *
	 */
	public static void addRPrShdStyle(RPr rpr, STShd shdtype, String color) {
		if (shdtype != null) {
		  CTShd shd = new CTShd();
	    shd.setVal(shdtype);
	    shd.setColor("auto");
	    shd.setFill(color);
	    rpr.setShd(shd);
	  }
	}

	/**
	 * 功能描述：获取段落属性对象
	 * @param paragraph
	 * @return
	 */
	public static PPr getPPr(P paragraph){
		PPr pprop = paragraph.getPPr();
		if (pprop == null) {
			pprop = new PPr();
			paragraph.setPPr(pprop);
		}
		return pprop;
	}

	public static Ind createInd(String firstLine , String firstLineChars){
		Ind  ind = new Ind();
		if(firstLine != null && firstLine.length() > 0){
			ind.setFirstLine(new BigInteger(firstLine));
		}
		if(firstLineChars != null && firstLineChars.length() > 0){
			ind.setFirstLineChars(new BigInteger(firstLineChars));
		}
		return ind;
	}

	/**
	 * 功能描述：获取间距对象
	 * @param ppr
	 * @return
	 */
	public static Spacing getSpacing(PPr ppr){
		Spacing spacing = ppr.getSpacing();
		if (spacing == null) {
			spacing = new Spacing();
			ppr.setSpacing(spacing);
		}
		return spacing;
	}


	/**
	* 功能描述：设置段落水平对齐方式
	* @param paragraph 段落对象
	* @param hAlign 对齐方式  JcEnumeration
	*/
	public static void setParaJcAlign(P paragraph, JcEnumeration hAlign) {
		if (hAlign != null) {
			PPr ppr = getPPr(paragraph);
			Jc align = new Jc();
			align.setVal(hAlign);
			ppr.setJc(align);
		}
	}

	/**
	 * 功能描述：设置段落边框样式
	 * @param p
	 * @param topBorder
	 * @param bottomBorder
	 * @param leftBorder
	 * @param rightBorder
	 */
	 public static void setParaBorders(P p, CTBorder topBorder, CTBorder bottomBorder, CTBorder leftBorder, CTBorder rightBorder) {
	    PPr ppr = getPPr(p);
	    PBdr pBdr = new PBdr();
	    if (topBorder != null) {
	      pBdr.setTop(topBorder);
	    }
	    if (bottomBorder != null) {
	      pBdr.setBottom(bottomBorder);
	    }
	    if (leftBorder != null) {
	      pBdr.setLeft(leftBorder);
	    }
	    if (rightBorder != null) {
	   	 pBdr.setRight(rightBorder);
	    }
	    ppr.setPBdr(pBdr);
	 }
		/**
		 * 添加图表
		 * @param wordMLPackage
		 * @param xlsxpath excel路径
		 * @param chartindex 图表序列
		 * @param headlist 表头
		 * @param datalist 数据
		 * @return
		 * @throws InvalidFormatException
		 * @throws FileNotFoundException
		 */
		/*public static P addChartToWord(WordprocessingMLPackage wordMLPackage, String xlsxpath, int chartindex, ArrayList<String> headlist, ArrayList<ArrayList<String>> datalist, String majorTitle, String minorTitle)
			throws InvalidFormatException, FileNotFoundException {
			Chart chart = DocxChart.createChart(wordMLPackage, xlsxpath, chartindex, headlist, datalist, majorTitle, minorTitle);
			return addChartToWord(wordMLPackage, chartindex, chart);
		}*/

		/**
		 * 添加图表
		 * @param wordMLPackage
		 * @param chartindex 图表序列
		 * @param chart 图表
		 * @return
		 * @throws InvalidFormatException
		 */
		public static P addChartToWord(WordprocessingMLPackage wordMLPackage, int chartindex, Chart chart) throws InvalidFormatException {
			String relid = wordMLPackage.getMainDocumentPart().addTargetPart(chart).getId();
			ObjectFactory factory = new ObjectFactory();
			Drawing drawing = createdocxChart(relid, chartindex);
			R run = factory.createR();
			run.getContent().add(drawing);
			P p = factory.createP();
			p.getContent().add(run);
			wordMLPackage.getMainDocumentPart().addObject(p);
			return p;
		}
		/**
		 * 创建文档图
		 * @param rid
		 * @param chartid
		 * @return
		 */
		public static Drawing createdocxChart(String rid, long chartid) {
			org.docx4j.dml.ObjectFactory dmlObjectFactory = new org.docx4j.dml.ObjectFactory();
			org.docx4j.dml.chart.ObjectFactory dmlchartObjectFactory = new org.docx4j.dml.chart.ObjectFactory();
			// Create object for chart (wrapped in JAXBElement)
			CTRelId relid = new CTRelId();
			relid.setId(rid);
			JAXBElement<CTRelId> relidWrapped = dmlchartObjectFactory.createChart(relid);
			//Create object for graphicData
			GraphicData graphicdata = dmlObjectFactory.createGraphicData();
			graphicdata.setUri("http://schemas.openxmlformats.org/drawingml/2006/chart");
			graphicdata.getAny().add(relidWrapped);
			//Create object for graphic
			Graphic graphic = dmlObjectFactory.createGraphic();
			graphic.setGraphicData(graphicdata);
			//Create object for cNvGraphicFramePr
			CTNonVisualGraphicFrameProperties nonvisualgraphicframeproperties = dmlObjectFactory.createCTNonVisualGraphicFrameProperties();
			//Create object for docPr
			CTNonVisualDrawingProps nonvisualdrawingprops = dmlObjectFactory.createCTNonVisualDrawingProps();
			nonvisualdrawingprops.setDescr("");
			nonvisualdrawingprops.setName("Chart 1");
			nonvisualdrawingprops.setId(chartid);
			//Create object for extent 实际面积
			CTPositiveSize2D positivesize2d = dmlObjectFactory.createCTPositiveSize2D();
			positivesize2d.setCx(5274310);
			positivesize2d.setCy(3076575);
			//Create object for effectExtent
			CTEffectExtent effectextent = new CTEffectExtent();
			effectextent.setT(0);
			effectextent.setL(0);//19050
			effectextent.setR(0);
			effectextent.setB(0);
			// Create object for inline
			Inline inline = new Inline();
			inline.setCNvGraphicFramePr(nonvisualgraphicframeproperties);
			inline.setDocPr(nonvisualdrawingprops);
			inline.setExtent(positivesize2d);
			inline.setEffectExtent(effectextent);
			inline.setGraphic(graphic);
			inline.setDistT(new Long(0));
			inline.setDistB(new Long(0));
			inline.setDistL(new Long(0));
			inline.setDistR(new Long(0));
			Drawing drawing = new Drawing();
			drawing.getAnchorOrInline().add(inline);
			dmlObjectFactory = null;
			dmlchartObjectFactory = null;
			return drawing;
		}


	 /**
	  * 获取空边框
	  * @return
	  */
	 public static CTBorder getnullBorder(){
		  	CTBorder ctb = new CTBorder();
	   	BigInteger v = new BigInteger("0");
	   	ctb.setSpace(v);
	   	ctb.setSz(v);
	   	ctb.setVal(STBorder.NONE);
	   	return ctb;
	 }

	 /**
	  * 创建目录域
	  * @param wordPackage
	  * @return
	  */
	 public static List<Object> createTOC(WordprocessingMLPackage wordPackage){
			SdtBlock sdtblock = createTOCBlock();
		 	addObject(wordPackage, sdtblock, false);
		 	ObjectFactory factory = Context.getWmlObjectFactory();
			P p = factory.createP();
			PPr ppr = factory.createPPr();
			p.setPPr(ppr);
//			PPrBase.Ind pprbaseind = factory.createPPrBaseInd();
//			ppr.setInd(pprbaseind);
			PStyle pprbasepstyle = factory.createPPrBasePStyle();
			ppr.setPStyle(pprbasepstyle);
			pprbasepstyle.setVal(StyleConstains.TOCHeading);
			R r = factory.createR();
			p.getContent().add(r);
			RPr rpr = factory.createRPr();
			r.setRPr(rpr);
			CTLanguage language = factory.createCTLanguage();
			rpr.setLang(language);
			language.setVal("zh-CN");
			Text text = factory.createText();
			JAXBElement<Text> textWrapped = factory.createRT(text);
			r.getContent().add(textWrapped);
			text.setValue("目录");
//			ContentAccessor sdtcontentblock = (ContentAccessor) sdtblock.getSdtContent();
			List<Object> list = sdtblock.getSdtContent().getContent();
			list.add(p);
			return list;
	 }

	 /**
	  * 设置 目录域标题
	  * @param sdtblock
	  * @param title
	  */
	 public static void setTOCTitle(SdtBlock sdtblock, String title){
			List<Object> toclist = sdtblock.getSdtContent().getContent();
			P  p = (P) toclist.get(0);
			R r = (R) p.getContent().get(0);
			Text t = ((JAXBElement<Text>) r.getContent().get(0)).getValue();
			t.setValue(title);
	 }
	 /**
	  * 创建目录域
	  * @return
	  */
	 public static SdtBlock createTOCBlock(){
		ObjectFactory factory = Context.getWmlObjectFactory();
		SdtBlock sdtblock = factory.createSdtBlock();
		SdtPr sdtpr = new SdtPr();
		sdtblock.setSdtPr(sdtpr);
		BigInteger sdtprid = sdtpr.setId();
//		Id id = factory.createId();
//		id.setVal( sdtprid );
//		sdtpr.getRPrOrAliasOrLock().add(id);

		//docPartObj
		CTSdtDocPart docPartObj = new CTSdtDocPart();
		DocPartGallery dpg = new DocPartGallery();
		dpg.setVal("Table of Contents");
		docPartObj.setDocPartGallery(dpg);
		BooleanDefaultTrue booleandefaulttrue = factory.createBooleanDefaultTrue();
		docPartObj.setDocPartUnique(booleandefaulttrue);
		sdtpr.getRPrOrAliasOrLock().add(factory.createSdtPrDocPartObj(docPartObj));

		//endpr
		CTSdtEndPr endpr = new CTSdtEndPr();
		sdtblock.setSdtEndPr(endpr);
		RPr endrpr = factory.createRPr();
		CTLanguage endlanguage = factory.createCTLanguage();
		endlanguage.setVal("zh-CN");
		endrpr.setLang(endlanguage);
		endpr.getRPr().add(endrpr);

		SdtContentBlock sdtcontentblock = factory.createSdtContentBlock();
		sdtblock.setSdtContent(sdtcontentblock);
		return sdtblock;
	 }

	 /**
	  * 手动加目录
	  * @param list
	  * @param content
	  * @param level
	  * @return
	  */
	 public static String addTOCCell(List<Object> list, String content , int level){
		String mark = System.currentTimeMillis() + "";
		mark = "_Toc" + mark.substring(4);//链接标记
		String pag = "0";//页码，无法确定默认为0
		String tocStyle = "TOC" + level;
		P p =creatHyperlinkToc(content, tocStyle, mark, pag);
		list.add(p);
		return mark;
	 };


	 public static P creatHyperlinkToc(String content , String tocStyle, String mark, String pag){

		 ObjectFactory wmlObjectFactory = Context.getWmlObjectFactory();
			// Create object for p
			P tocp = wmlObjectFactory.createP();
			// Create object for pPr
			PPr ppr = wmlObjectFactory.createPPr();
			tocp.setPPr(ppr);
			// Create object for rPr
			ParaRPr pararpr = wmlObjectFactory.createParaRPr();
			ppr.setRPr(pararpr);
			// Create object for noProof
			BooleanDefaultTrue booleandefaulttrue = wmlObjectFactory.createBooleanDefaultTrue();
			pararpr.setNoProof(booleandefaulttrue);
			// Create object for sz
			HpsMeasure hpsmeasure01 = wmlObjectFactory.createHpsMeasure();
			hpsmeasure01.setVal(BigInteger.valueOf(21));
			pararpr.setSz(hpsmeasure01);
			// Create object for kern
			HpsMeasure hpsmeasure02 = wmlObjectFactory.createHpsMeasure();
			hpsmeasure02.setVal(BigInteger.valueOf(2));
			pararpr.setKern(hpsmeasure02);

			// Create object for pStyle
			PStyle pprbasepstyle2 = wmlObjectFactory.createPPrBasePStyle();
			ppr.setPStyle(pprbasepstyle2);
			pprbasepstyle2.setVal(tocStyle);
			// Create object for tabs
			Tabs tabs = wmlObjectFactory.createTabs();
			ppr.setTabs(tabs);
			// Create object for tab
			CTTabStop tabstop = wmlObjectFactory.createCTTabStop();
			tabs.getTab().add(tabstop);
			tabstop.setPos(BigInteger.valueOf(8296));
			tabstop.setVal(STTabJc.RIGHT);
			tabstop.setLeader(STTabTlc.DOT);


//			R r3 = addFieldBegin(tocp.getContent());
//			RPr rpr3 = wmlObjectFactory.createRPr();
//			r3.setRPr(rpr3);
//			rpr3.setNoProof(booleandefaulttrue);

//	            // Create object for r
//	            R r3 = wmlObjectFactory.createR();
//	            p3.getContent().add( r3);
//	            // Create object for instrText (wrapped in JAXBElement)
//	            Text text2 = wmlObjectFactory.createText();
//	            JAXBElement<org.docx4j.wml.Text> textWrapped2 = wmlObjectFactory.createRInstrText(text2);
//	            r3.getContent().add( textWrapped2);
//	            text2.setValue( " TOC \\o \"1-3\" \\h \\z \\u ");
//	            text2.setSpace( "preserve");

//			R r4 = addFieldSeparate(tocp.getContent());
//			RPr rpr4 = wmlObjectFactory.createRPr();
//			r4.setRPr(rpr4);
//			rpr4.setNoProof(booleandefaulttrue);

			// 创建链接
			P.Hyperlink phyperlink = wmlObjectFactory.createPHyperlink();
			//设置锚
			phyperlink.setAnchor(mark);

			R r5 = wmlObjectFactory.createR();
			RPr rpr5 = wmlObjectFactory.createRPr();
			r5.setRPr(rpr5);
			BooleanDefaultTrue booleandefaulttrue5 = wmlObjectFactory.createBooleanDefaultTrue();
			rpr5.setNoProof(booleandefaulttrue5);
			RStyle rstyle = wmlObjectFactory.createRStyle();
			rpr5.setRStyle(rstyle);
			//链接设置目录样式
			rstyle.setVal("Hyperlink");
			// Create object for t (wrapped in JAXBElement)
			Text text3 = wmlObjectFactory.createText();
			JAXBElement<Text> textWrapped3 = wmlObjectFactory.createRT(text3);
			r5.getContent().add(textWrapped3);
			text3.setValue(content);
			phyperlink.getContent().add(r5);

			R r6 = wmlObjectFactory.createR();
			setTOCRpr(r6);
			R.Tab rtab = wmlObjectFactory.createRTab();
			JAXBElement<R.Tab> rtabWrapped = wmlObjectFactory.createRTab(rtab);
			r6.getContent().add(rtabWrapped);
			phyperlink.getContent().add(r6);

//			R r7 = addFieldBegin(phyperlink.getContent());
			R r7 = getFieldBegin(false);
			phyperlink.getContent().add(r7);
			setTOCRpr(r7);

			R r8 = addFieldContent(phyperlink.getContent(), " PAGEREF " + mark + " \\h ");
			setTOCRpr(r8);

			R r9 = wmlObjectFactory.createR();
			setTOCRpr(r9);
			phyperlink.getContent().add(r9);

//			R r10 = addFieldSeparate(phyperlink.getContent());
			R r10 = getFieldSeparate();
			phyperlink.getContent().add(r10);
			setTOCRpr(r10);

			R r11 = wmlObjectFactory.createR();
			setTOCRpr(r11);

			Text text5 = wmlObjectFactory.createText();
			text5.setValue(pag);
			JAXBElement<Text> textWrapped5 = wmlObjectFactory.createRT(text5);
			r11.getContent().add(textWrapped5);
			phyperlink.getContent().add(r11);

//			R r12 = addFieldEnd(phyperlink.getContent());
			R r12 = getFieldEnd();
			phyperlink.getContent().add(r12);
			setTOCRpr(r12);

			JAXBElement<P.Hyperlink> phyperlinkWrapped = wmlObjectFactory.createPHyperlink(phyperlink);
			tocp.getContent().add(phyperlinkWrapped);
			return tocp;

	 }

	 private static void setTOCRpr(R r){
		ObjectFactory factory = Context.getWmlObjectFactory();
		RPr rpr = factory.createRPr();
		r.setRPr(rpr);
		BooleanDefaultTrue booleandefaulttrue6 = factory.createBooleanDefaultTrue();
		rpr.setNoProof(booleandefaulttrue6);
		BooleanDefaultTrue booleandefaulttrue7 = factory.createBooleanDefaultTrue();
		rpr.setWebHidden(booleandefaulttrue7);
	 }


//	 public static void addTableOfContentAuto(WordprocessingMLPackage wordPackage, int level){
//		 ObjectFactory factory = Context.getWmlObjectFactory();
//		 CTSdtCell sdt = factory.createCTSdtCell();
//		 SdtPr sdtpr = new SdtPr();
//		 sdt.setSdtPr(sdtpr);
//		 BigInteger sdtprid = sdtpr.setId();
//		 //docPartObj
//		 CTSdtDocPart docPartObj = new CTSdtDocPart();
//		 DocPartGallery dpg = new DocPartGallery();
//		 dpg.setVal("Table of Contents");
//		 docPartObj.setDocPartGallery(dpg);
//		 docPartObj.setDocPartUnique(null);
//		 sdtpr.getRPrOrAliasOrLock().add(factory.createSdtPrDocPartObj(docPartObj));
//		 //endpr
//		 CTSdtEndPr endpr = new CTSdtEndPr();
//		 sdt.setSdtEndPr(endpr);
//		 RPr rpr = getRPr(null, "auto", "22", null, null);
//		 endpr.getRPr().add(rpr);
//		 RFonts rf = factory.createRFonts();
//		 rpr.setRFonts(rf);
//		 rf.setAsciiTheme(STTheme.MINOR_H_ANSI);
//		 rf.setEastAsiaTheme(STTheme.MINOR_H_ANSI);
//		 rf.setHAnsiTheme(STTheme.MINOR_H_ANSI);
//		 rf.setCstheme(STTheme.MINOR_BIDI);
//		 BooleanDefaultTrue b = new BooleanDefaultTrue();
//		 b.setVal(false);
//		 rpr.setB(b);
//		 rpr.setBCs(b);
//		 //sdtcontent
//		 CTSdtContentCell sdtcontent = new CTSdtContentCell();
//		 sdt.setSdtContent(sdtcontent);
//
//		 P p = createStyledParagraphOfText(wordPackage, "TOCHeading", "目录");
//		 R r = (R)p.getContent().get(0);
////		 r.setRPr(getjsnkTabelofcontent());
////		 setParaJcAlign(p, JcEnumeration.CENTER);
////		 setParaBorders(p, null,getnullBorder(),null,null);
//		 sdtcontent.getContent().add(p);
//
//		 P paragraph = factory.createP();
//		 addFieldBegin(paragraph);
//	    R run = factory.createR();
//	    Text txt = new Text();
//	    txt.setSpace("preserve");
//	    txt.setValue("TOC \\o \"1-"+ level +"\" \\h \\z \\u");
//	    run.getContent().add(factory.createRInstrText(txt));
//	    paragraph.getContent().add(run);
////		 addTableOfContentField(paragraph);
//	    addFieldSeparate(paragraph);
//	    addFieldEnd(paragraph);
//
//		 sdtcontent.getContent().add(paragraph);
//
//		 addObject(wordPackage, sdt, false);
//	 }

   /**
    *  将目录表添加到文档.
    *
    *  首先我们创建段落. 然后添加标记域开始的指示符, 然后添加域内容(真正的目录表), 接着添加域
    *  结束的指示符. 最后将段落添加到给定文档的JAXB元素中.
    *
    *  @param wordPackage
    */
	public static void addTableOfContent(WordprocessingMLPackage wordPackage) {
		addTableOfContent(wordPackage, 3);
	}

	/**
	 *
	 * @param wordPackage
	 * @param level 目录显示层级
	 */
   public static void addTableOfContent(WordprocessingMLPackage wordPackage,int level) {
   	ObjectFactory factory = Context.getWmlObjectFactory();
   	P p = addStyledParagraphOfText(wordPackage, "TOCHeading", "目录");
      P paragraph = factory.createP();
      addAotuTOC(paragraph, level, false);
      addObject(wordPackage, paragraph, false);
   }

   /**
    * 是否手写
    * @param paragraph
    * @param level
    * @param ishandwrit
    */
   public static void addAotuTOC(P paragraph, int level, boolean ishandwrit){
   	R beginrun = getFieldBegin(!ishandwrit);
   	paragraph.getContent().add(0, beginrun);
   	addTableOfContentField(paragraph, level);

		R separaterun = getFieldSeparate();
   	paragraph.getContent().add(2, separaterun);

//      if(ishandwrit){
//      }else{
//   		R endrun = getFieldEnd();
//      	paragraph.getContent().add(3, endrun);
//   	}
   }

   /**
    *  将Word用于创建目录表的域添加到段落中.
    *  首先创建一个可运行块和一个文本. 然后指出文本中所有的空格都被保护并给TOC域设置值. 这个域定义
	 *  需要一些参数, 确切定义可以在Office Open XML标准的§17.16.5.58找到, 这种情况我们指定所有
	 *  段落使用1-3级别的标题来格式化(\0 "1-3"). 我们同时指定所有的实体作为超链接(\h), 而且在Web
	 *  视图中隐藏标签和页码(\z), 我们要使用段落大纲级别应用(\\u).
    *
    * @param paragraph
    */
   public static void addTableOfContentField(P paragraph, int level) {
   	String content = " TOC \\o \"1-" + level + "\" \\h \\z \\u ";
   	R run = getFieldContent(content);
   	List<Object> list = paragraph.getContent();
   	list.add(1, run);
   }

   /**
    *
    * @param paragraph
    * @param content
    */
   public static void addFieldContent(P paragraph, String content){
   	addFieldContent(paragraph.getContent(), content);
   }

   private static R addFieldContent(List<Object> list, String content){
   	R run = getFieldContent(content);
      list.add(run);
      return run;
   }

	private static R getFieldContent(String content) {
		ObjectFactory factory = Context.getWmlObjectFactory();
      R run = factory.createR();
      Text txt = getText(content, "preserve");
      run.getContent().add(factory.createRInstrText(txt));
		return run;
	}

   /**
    *  每个域都需要用复杂的域字符来确定界限. 本方法向给定段落添加在真正域之前的界定符.
    *
    *  再一次以创建一个可运行块开始, 然后创建一个域字符来标记域的起始并标记域是'脏的'因为我们想要
    *  在整个文档生成之后进行内容更新.
    *  最后将域字符转换成JAXB元素并将其添加到可运行块, 然后将可运行块添加到段落中.
    *
    *  @param paragraph
    */
   public static void addFieldBegin(P paragraph, boolean dirty) {
   	paragraph.getContent().add(getFieldBegin(dirty));
   }

	private static R getFieldBegin(boolean dirty) {
		ObjectFactory factory = Context.getWmlObjectFactory();
      R run = factory.createR();
      FldChar fldchar = factory.createFldChar();
      fldchar.setFldCharType(STFldCharType.BEGIN);
      if(dirty){
      	fldchar.setDirty(dirty);
      }
      run.getContent().add(getWrappedFldChar(fldchar));
		return run;
	}

   /**
    *  每个域都需要用复杂的域字符来确定界限. 本方法向给定段落添加在真正域之后的界定符.
    *
    *  跟前面一样, 从创建可运行块开始, 然后创建域字符标记域的结束, 最后将域字符转换成JAXB元素并
    *  将其添加到可运行块, 可运行块再添加到段落中.
    *
    *  @param paragraph
    */
   public static void addFieldEnd(P paragraph) {
   	paragraph.getContent().add(getFieldEnd());
   }

	private static R getFieldEnd() {
		ObjectFactory factory = Context.getWmlObjectFactory();
      R run = factory.createR();
      FldChar fldcharend = factory.createFldChar();
      fldcharend.setFldCharType(STFldCharType.END);
      run.getContent().add(getWrappedFldChar(fldcharend));
		return run;
	}
   /**
    * 本方法向给定段落添加在真正域分隔的界定符.
    *
    * @param paragraph
    */
   private static void addFieldSeparate(P paragraph) {
   	paragraph.getContent().add(getFieldSeparate());
   }

   private static R getFieldSeparate(){
   	ObjectFactory factory = Context.getWmlObjectFactory();
      R run = factory.createR();
      FldChar fldcharend = factory.createFldChar();
      fldcharend.setFldCharType(STFldCharType.SEPARATE);
      run.getContent().add(getWrappedFldChar(fldcharend));
   	return run;
   }

   /**
    *  创建包含给定复杂域字符的JAXBElement的便利方法.
    *
    *  @param fldchar
    *  @return
    */
	private static JAXBElement getWrappedFldChar(FldChar fldchar) {
       return new JAXBElement(new QName(Namespaces.NS_WORD12, "fldChar"), FldChar.class, fldchar);
   }
	/*********************  图片  ******* **********/

	/**
	 * 功能描述：添加图片
	 * @param wordPackage
	 * @param imagePath
	 * @return 返回该图片的段落
	 * @throws Exception
	 */
	public static P addImage( WordprocessingMLPackage wordPackage, String imagePath ) throws Exception {
		P p = getPImage(wordPackage, imagePath);
		addObject(wordPackage, p, false);
		return p;
	}
	/**
	 * 功能描述：创建图片信息
	 * @param wordPackage
	 * @param imagePath
	 * @return 返回该图片的段落
	 * @throws Exception
	 */
	public static P getPImage( WordprocessingMLPackage wordPackage, String imagePath ) throws Exception {
		return getPImage(wordPackage, getImageByte(imagePath), "Picture "+System.currentTimeMillis(), "", 0, 1);
	}
	/**
	 * 功能描述：创建图片信息
	 * @param wordPackage  文档处理包对象
	 * @param bytes        字节数组
	 * @param filenameHint 图片名称
	 * @param altText      提示信息
	 * @param id1          id1
	 * @param id2          id2
	 * @return 返回该图片的段落
	 * @throws Exception
	 * @author cnbi
	 */
	public static P getPImage( WordprocessingMLPackage wordPackage,byte[] bytes,
		String filenameHint, String altText, int id1, int id2) throws Exception {
      BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordPackage, bytes);
      Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, false);
		ObjectFactory factory = new ObjectFactory();
		P  p = factory.createP();
		R  run = factory.createR();
		p.getContent().add(run);
		Drawing drawing = factory.createDrawing();
		run.getContent().add(drawing);
		drawing.getAnchorOrInline().add(inline);
		PPr pPr = p.getPPr();
		if(pPr == null){
			pPr = factory.createPPr();
		}
		Jc jc = pPr.getJc();
		if(jc == null){
			jc = new Jc();
		}
		jc.setVal(JcEnumeration.CENTER);
		pPr.setJc(jc);
		p.setPPr(pPr);
		return p;
	}
	/**
	 * 功能描述：获取图片的btye[]
	 * @param path
	 * @return
	 * @throws IOException
	 * @throws NullPointerException
	 */
	public static byte[] getImageByte(String path ) throws IOException,NullPointerException{
			File file = new File(path);
			long length = file.length();
			// You cannot create an array using a long type.
			// It needs to be an int type.
			if (length > Integer.MAX_VALUE) {

				throw new NumberFormatException("greater Integer.MAX_VALUE");
			}
			byte[] bytes = new byte[(int)length];
			int offset = 0;
			int numRead = 0;
			java.io.InputStream is = null;
			try {
			// Our utility method wants that as a byte array
				is = new FileInputStream(file );
				while (offset < bytes.length
					&& (numRead=is.read(bytes, offset, bytes.length-offset)) >= 0) {
					offset += numRead;
				}
			} catch (IOException e) {
				e.printStackTrace();

				throw e;
			}finally{
				if(is != null){
					try {
						is.close();
					} catch (IOException e) {
					}
				}
			}
//			// Ensure all the bytes have been read in
//			if (offset < bytes.length) {
//	         System.out.println("Could not completely read file "+file.getName());
//			}
			return bytes;

	}

	/*********************  文档表格  ******* **********/
//  循环文档表格
//	public static void forTbl(Tbl tbl, STShd shdtype ){
//		List<Object> rowList = tbl.getContent();
//		for(int row=0, rowsize=rowList.size(); row<rowsize; row++){
//			Tr tr = (Tr) XmlUtils.unwrap(rowList.get(row));
//			List<Object> colList = tr.getContent();
//			for(int col=0, colsize=colList.size(); col<colsize; col++){
//				Tc tc = (Tc) XmlUtils.unwrap(colList.get(col));
//			}
//		}
//	}


	/**
	 * 功能描述：添加文档表格
	 * @param wordPackage 文档处理包对象
	 * @param rows 表格行数
	 * @param cols 表格列数
	 * @return  返回表格对象
	 * @throws Exception
	 */
	public static Tbl addTable(WordprocessingMLPackage wordPackage , int rows , int cols, int[] width) throws Exception{
		Tbl tb = createTable(wordPackage, rows, cols, width);
		List<Object> list = wordPackage.getMainDocumentPart().getContent();
		if(list.size() > 0){
			Object obj = list.get(list.size()-1);
			if(obj instanceof Tbl){
				addBodyParagraphOfText(wordPackage, "");
			}
		}
		addObject(wordPackage, tb, false);
		return tb;
	}
	/**
	 * 功能描述：创建文档表格，表格的默认宽度为：9328，表格样式：dxa，对齐方式：居中
	 * @param wordPackage 文档处理包对象
	 * @param rows        表格行数
	 * @param cols        表格列数
	 * @return            返回值：返回表格对象
	 * @throws Exception
	 * @author cnbi
	 */
	public static Tbl createTable(WordprocessingMLPackage wordPackage , int rows , int cols, int[] width) throws Exception{
		return createTable(wordPackage, rows, cols, width, "dxa", JcEnumeration.CENTER);
	}

	/**
	 * 功能描述：创建文档表格，表格的默认宽度为：表格样式：dxa，对齐方式：居中
	 * @param wordPackage 文档处理包对象
	 * @param rows        表格行数
	 * @param cols        表格列数
	 * @param tableWidth  表格宽度
	 * @return            返回值：返回表格对象
	 * @throws Exception
	 * @author cnbi
	 */
	/*public static Tbl createTable(WordprocessingMLPackage wordPackage , int rows , int cols , String tableWidth) throws Exception{
		return createTable(wordPackage, rows, cols, tableWidth , "dxa", JcEnumeration.CENTER);
	}*/

	/**
	 * 功能描述：创建文档表格，全部表格边框都为1
	 * @param wordPackage  文档处理包对象
	 * @param rows         表格行数
	 * @param cols         表格列数
	 * @param width   表格的宽度
	 * @param style        表格的样式    按内容 pct 按窗口 auto 固定列 dxa
	 * @param jcEnumerationVal 表格的对齐方式
	 * @return             返回值：返回表格对象
	 * @throws Exception
	 * @author cnbi
	 */
	public static Tbl createTable(WordprocessingMLPackage wordPackage , int rows , int cols , int[] width , String style , JcEnumeration jcEnumerationVal)throws Exception{
		int writableWidthTwips = getWritableWidth(wordPackage);
		System.out.println("文档可用宽度==>"+writableWidthTwips);
		if(cols == 0){
			cols = 1;
		}
		int cellWidth = new Double(Math.floor( (writableWidthTwips/cols ))).intValue();
		int[] widths = new int[cols];
		int len = 0;
		for(int i = 0 ; i < cols ; i++){
		   int a = width[i] * 50;
		   len += a;
		   widths[i] = a;
		}
		if(len >= writableWidthTwips)
		    for(int i = 0 ; i < cols ; i++){
			widths[i] =  widths[i] > cellWidth? cellWidth : widths[i];
		    }
		Tbl tbl =  createBorderTable(rows, cols, widths);
		TblPr tblpr = getTblPr(tbl);
		TblWidth tblw = new TblWidth();
		tblw.setType("pct");
		tblw.setW(new BigInteger("5000"));
		tblpr.setTblW(tblw);
		return tbl;
	}

	/**
	 * 功能描述：创建文档表格，全部表格边框都为1
	 * @param rows    行数
	 * @param cols    列数
	 * @param widths  每列的宽度
	 * @return   返回值：返回表格对象
	 * @author cnbi
	 */
	public static Tbl createBorderTable(int rows, int cols, int[] widths) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		Tbl tbl = factory.createTbl();
		// w:tblPr
		StringBuffer tblSb = new StringBuffer();
		tblSb.append("<w:tblPr ").append(Namespaces.W_NAMESPACE_DECLARATION).append(">");
		tblSb.append("<w:tblStyle w:val=\"TableGrid\"/>");
		tblSb.append("<w:tblW w:w=\"0\" w:type=\"auto\"/>");
		tblSb.append("<w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
		tblSb.append("<w:left w:val=\"single\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>");
		tblSb.append("<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
		tblSb.append("<w:right w:val=\"single\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>");
		tblSb.append("</w:tblBorders>");
		tblSb.append("<w:tblLook w:val=\"04A0\"/>");
		tblSb.append("</w:tblPr>");
		TblPr tblPr = null;
		try {
			tblPr = (TblPr) XmlUtils.unmarshalString(tblSb.toString());
		} catch (JAXBException e) {
			e.printStackTrace();
		}
		tbl.setTblPr(tblPr);
		if (tblPr != null) {
			Jc jc = factory.createJc();
			//单元格居中对齐
			jc.setVal(JcEnumeration.CENTER);
			tblPr.setJc(jc);
			CTTblLayoutType tbll = factory.createCTTblLayoutType();
			// 固定列宽
			tbll.setType(STTblLayoutType.AUTOFIT);
			tblPr.setTblLayout(tbll);
		}
		// <w:tblGrid><w:gridCol w:w="4788"/>
		TblGrid tblGrid = factory.createTblGrid();
		tbl.setTblGrid(tblGrid);
		// Add required <w:gridCol w:w="4788"/>
		for (int i = 1; i <= cols; i++) {
			TblGridCol gridCol = factory.createTblGridCol();
			gridCol.setW(BigInteger.valueOf(widths[i - 1]));
			tblGrid.getGridCol().add(gridCol);
		}
		// Now the rows
		for (int j = 1; j <= rows; j++) {
			Tr tr = factory.createTr();
			tbl.getContent().add(tr);
			setTrHight(tr, 6f, null);
			// The cells
			for (int i = 1; i <= cols; i++) {
				Tc tc = factory.createTc();
				tr.getContent().add(tc);
				TcPr tcPr = getTcPr(tc);
				// <w:tcW w:w="4788" w:type="dxa"/>
				TblWidth cellWidth = factory.createTblWidth();
				tcPr.setTcW(cellWidth);
				cellWidth.setW(BigInteger.valueOf(0/*widths[i - 1]*/));
				cellWidth.setType("auto");
				P p = factory.createP();
				PPr ppr  = getPPr(p);
				Spacing spacing = getSpacing(ppr);
				spacing.setLine(new BigInteger("240"));
				spacing.setLineRule(STLineSpacingRule.AUTO);
				tc.getContent().add(p);
			}
		}
		return tbl;
	}

	/**
	 * 功能描述：设置表给单元格内容对齐方式
	 * @param tbl 表格对象
	 * @param startRow 起始行
	 * @param startCol 起始列
	 * @param rows 行数
	 * @param cols 列数
	 * @param hAlign 表格水平对齐方式  JcEnumeration
	 * @param vTcAlign
	 */
	public static void setTablecellContentStyle(Tbl tbl,int startRow,int startCol,int rows, int cols,JcEnumeration hAlign,STVerticalJc vTcAlign ){
		List rowList = tbl.getContent();
		int row = startRow, endrow = startRow + rows-1;
		for(;row <= endrow;row++){
			int col  = startCol, endcol=startCol+cols-1;
			Tr tr = (Tr) XmlUtils.unwrap(rowList.get(row));
			List colList = tr.getContent();
			for(;col <= endcol;col++){
				Tc tc = (Tc) XmlUtils.unwrap(colList.get(col));
				List<Object> plist =  tc.getContent();
				int size = plist.size();
				for(int i = 0;i<size;i++){
					setParaJcAlign((P) plist.get(i),hAlign);
				}
				setTcVAlign(tc, vTcAlign);
			}
		}
	}

	/**
	 *
	 * @param magelist 合并单元格集
	 * @param tbl 表格对象
	 */
	public static void mageCell(ArrayList<int[]> magelist, Tbl tbl) {
//		合并单元格前需对单元格排序
//		按先从下至上后从右向左的顺序合并
		if (magelist.size() > 1) {
				// System.out.println("---->"+heandermagelist.get(3)[1]);
				Collections.sort(magelist, new Comparator<int[]>() {
					public int compare(int[] tree0, int[] tree1) {
						int desc = 0;
						if(tree0[0] == tree1[0]){
							desc = tree0[1] == tree1[1] ? 0 : tree0[1] > tree1[1] ? -1 : 1;
						}else if( tree0[0] > tree1[0] ){
							desc = -1;
						}else{
							desc = 1;
						}
						return desc;
					}
				});
		}
		for (int i = 0, size = magelist.size(); i < size; i++) {
			int[] merge = magelist.get(i);
			try {
				DOCUtils.setCellMerge(tbl, merge[0], merge[2], merge[1], merge[3]);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}// C2D69B C6D9F1

	}

	/**
	 * 功能描述：合并单元格</BR>
	 *			表示合并第startRow（开始行）行中的第startCol（开始列）列到（startCol + colSpan - 1）列 </BR>
	 *			表示合并第startCol（开始列）行中的第startRow（开始行）列到（startRow + rowSpan - 1）行
	 * @param tbl          表格对象
	 * @param startRow    开始行
	 * @param rowSpan     合并的行数，大于1才表示合并
	 * @param startCol    开始列
	 * @param colSpan     合并的列数，大于1才表示合并
	 * @author cnbi
	 */
	public static void setCellMerge(Tbl tbl , int startRow , int rowSpan , int startCol , int colSpan){
		ObjectFactory factory = Context.getWmlObjectFactory();
		List<Object> rowList = tbl.getContent();
		int row = startRow,endrow=startRow + rowSpan - 1;
		for(;row<=endrow;row++){
			Tr tr = (Tr) XmlUtils.unwrap(rowList.get(row));
			List<Object> colList = tr.getContent();
			int col = startCol,endcol = startCol + colSpan - 1;
			for(;endcol >= col;endcol--){
				if(endcol == startCol){
					Tc tc = (Tc) XmlUtils.unwrap(colList.get(endcol));
					TcPr tcpr =tc.getTcPr();
					if(colSpan>1){
						GridSpan gridspan = factory.createTcPrInnerGridSpan();
						BigInteger span = new BigInteger(colSpan+"");
						gridspan.setVal(span);
						tcpr.setGridSpan(gridspan);
						BigInteger tcww = tcpr.getTcW().getW().multiply(span);
						tcpr.getTcW().setW(tcww);
					}
					if(rowSpan>1 ){
						VMerge vm =  factory.createTcPrInnerVMerge();
						if(row == startRow){
							vm.setVal("restart");
						}
						tcpr.setVMerge(vm);
					}
				}else{
					colList.remove(endcol);
				}
			}
		}
	}

	/**
	 * 功能描述：填充表格内容，固定表头，表头宋体加粗，小五号，表格内容宋体，小五号，表格居中对齐</BR>
	 * 			其中表格数据跟表头数据结构要一致，适用于简单的n行m列的普通表格
	 * @param wordPackage  文档处理包对象
	 * @param tbl          表格对象
	 * @param dataList     表格数据
	 * @param titleList    表头数据，假如不需要表头信息，则只要传进null即可
	 * @author cnbi
	 */
	public static void fillTableData(WordprocessingMLPackage wordPackage , Tbl tbl , List dataList , List titleList){
		fillTableData(wordPackage, tbl, dataList, titleList, JcEnumeration.CENTER, JcEnumeration.CENTER);
	}

	/**
	 * 功能描述：填充表格内容，表头宋体加粗，小五号，表格内容宋体，小五号，表格居中对齐</BR>
	 * 			其中表格数据跟表头数据结构要一致，适用于简单的n行m列的普通表格
	 * @param wordPackage  文档处理包对象
	 * @param tbl          表格对象
	 * @param dataList     表格数据
	 * @param titleList    表头数据，假如不需要表头信息，则只要传进null即可
	 * @param tjcEnumeration 表头对齐方式
	 * @param jcEnumeration 表格水平对齐方式
	 * @author cnbi
	 */
	public static void fillTableData(WordprocessingMLPackage wordPackage , Tbl tbl , List dataList , List titleList, JcEnumeration tjcEnumeration, JcEnumeration jcEnumeration ){
		fillTableData(wordPackage, tbl, dataList, titleList, false, DocxConstains.fontFamily, DocxConstains.fontSize, true, tjcEnumeration, DocxConstains.fontFamily, DocxConstains.fontSize, false, jcEnumeration);
	}

	/**
	 * 功能描述：填充表格内容
	 * @param wordPackage  文档处理包对象
	 * @param tbl          表格对象
	 * @param dataList     表格数据
	 * @param titleList    表头数据
	 * @param isFixedTitle 是否固定表头  ：office word中 表格属性：行：在各页顶端以标题形式重复出现
	 * @param tFontFamily  表头字体
	 * @param tFontSize    表头字体大小
	 * @param tIsBlod      表头是否加粗
	 * @param tJcEnumeration 表头对齐方式
	 * @param fontFamily   表格字体
	 * @param fontSize     表格字号
	 * @param isBlod       表格内容是否加粗
	 * @param jcEnumeration 表格水平对齐方式
	 * @author cnbi
	 */
	public static void fillTableData(WordprocessingMLPackage wordPackage , Tbl tbl , List dataList , List titleList , boolean isFixedTitle,String tFontFamily , String tFontSize , boolean tIsBlod , JcEnumeration tJcEnumeration,String fontFamily , String fontSize , boolean isBlod , JcEnumeration jcEnumeration){
		List rowList = tbl.getContent();
		//boolean flag = ((List)titleList.get(0)).get(0).equals("项目");//首创要求项目加粗斜体
		//整个表格的行数
		int rows = rowList.size();
		int tSize = 0;//猎头行数
		//表头
		if(titleList != null && titleList.size() > 0){
			tSize = titleList.size();
			for(int t = 0 ; t < tSize ; t++){

				Tr tr0 = (Tr) XmlUtils.unwrap(rowList.get(t));
				List colList = tr0.getContent();
//				Object[] tobj = (Object[]) titleList.get(t);
				List<Object> tobj = (List<Object>) titleList.get(t);
				for(int c = 0 , size=colList.size(), tsize = tobj.size(); c < size && c < tsize ; c++){
					Tc tc0 = (Tc) XmlUtils.unwrap(colList.get(c));
					//填充表头数据
					fillCellData(tc0, converObjToStr(tobj.get(c), ""), tFontFamily, tFontSize, tIsBlod, tJcEnumeration, null);

				}
				if(isFixedTitle){
					//设置固定表头
					fixedTitle(tr0);
				}
			}
		}
		int colsSize = 1;
		//表格数据，往掉表头之后的表格数据
		for(int r = tSize ; r < rows ; r++){
			Tr tr = (Tr) XmlUtils.unwrap(rowList.get(r));
//			Object[] objs = null;
			//假如表格内容不为空，则取出相应的数据进行填充
//			if(dataList != null && (dataList.size() >= (rows - tSize))){
//				objs = (Object[]) dataList.get(r - tSize);
//			}

			List colsList = tr.getContent();
			List<Object> objs = (List<Object>) dataList.get(r - tSize);
			//定义“合计”所在的行数 合计那一行加粗
			int rowNuOfCount = -1;
			//整个表格的列数
			colsSize = colsList.size();
			for(int i = 0 ,dsize = objs.size() ; i < colsSize && i < dsize ; i++){
				Tc tc = (Tc) XmlUtils.unwrap(colsList.get(i));
				//填充表格数据
				String value = converObjToStr(objs.get(i), "");
				JcEnumeration datajce =  Docx4jToolUtils.isNumb(value.replaceAll("%", "")) || "--".equals(value)? jcEnumeration: JcEnumeration.LEFT;

				if(i == 0) {
					if (value.contains("合计")){
						//合计所在的行
						rowNuOfCount = r;
						fillCellData(tc, value, fontFamily, fontSize, true, datajce, null);

					} else {
						fillCellData(tc, value, fontFamily, fontSize, isBlod, datajce, null);
					}

				}
				else {
					if (r == rowNuOfCount) {
						//合计所在行字体加粗
						fillCellData(tc, value, fontFamily, fontSize, true, datajce, null);
					} else {
						fillCellData(tc, value, fontFamily, fontSize, isBlod, datajce, null);
					}
				}
//				if(objs != null){
//					fillCellData(tc, converObjToStr(objs[i], ""), fontFamily, fontSize, isBlod, jcEnumeration);
//				}else{
//					fillCellData(tc, "", fontFamily, fontSize, isBlod, jcEnumeration);
//				}
			}
		}
	}

	/**
	 * 功能描述：固定表头
	 * @param tr 行对象
	 * @author cnbi
	 */
	public static void fixedTitle(Tr tr){
		ObjectFactory factory = Context.getWmlObjectFactory();
		BooleanDefaultTrue bdt = factory.createBooleanDefaultTrue();
		//表示固定表头
		bdt.setVal(true);
		TrPr trpr = tr.getTrPr();
		if(trpr == null){
			trpr = factory.createTrPr();
			tr.setTrPr(trpr);
		}
		trpr.getCnfStyleOrDivIdOrGridBefore().add(factory.createCTTrPrBaseTblHeader(bdt));
	}

	/**
	 * 功能描述：设置行高
	 * @param tr
	 * @param mm 毫米
	 * @param heightrule 行高规则
	 */
	public static void setTrHight(Tr tr, float mm,  STHeightRule heightrule){
		ObjectFactory factory = Context.getWmlObjectFactory();
		TrPr trpr = getTrPr(tr);
		CTHeight cth = new CTHeight();
		cth.setHRule(heightrule);
		cth.setVal(new BigInteger(UnitsOfMeasurement.mmToTwip(mm)+""));
		trpr.getCnfStyleOrDivIdOrGridBefore().add(factory.createCTTrPrBaseTrHeight(cth));
	}

	/**
	 * 功能描述：填充单元格内容
	 * @param tc          单元格对象
	 * @param data        内容
	 * @param fontFamily  字体
	 * @param fontSize    字号
	 * @param isBlod      是否加粗
	 * @param jcEnumeration 对齐方式
	 * @author cnbi
	 */
	public static void fillCellData(Tc tc , String data , String fontFamily , String fontSize , boolean isBlod , JcEnumeration jcEnumeration, Boolean flag){
		ObjectFactory factory = Context.getWmlObjectFactory();
		P p = (P) XmlUtils.unwrap(tc.getContent().get(0));
		//设置表格内容的对齐方式
		setParaJcAlign(p , jcEnumeration);
		Text t = factory.createText();
		t.setValue(data);
		R run = factory.createR();
		//设置表格内容字体样式
		//run.setRPr(getRPr(fontFamily, fontSize, isBlod));
		run.setRPr(getRPrForSCGFTable(fontFamily, isBlod, flag));
		//设置内容垂直居中
		setTcVAlign(tc, STVerticalJc.CENTER);
		run.getContent().add(t);

		p.getContent().add(run);

		PPr ppr = getPPr(p);
		ppr.setInd(createInd("0","0"));
	}

	/**
	 * 功能描述：取行配置对象
	 * @param tr
	 * @return
	 */
	public static TrPr getTrPr(Tr tr) {
		TrPr trpr = tr.getTrPr();
		if(trpr == null){
			trpr = new TrPr();
			tr.setTrPr(trpr);
		}
		return trpr;
	}

	/**
	 * 功能描述：取表格配置对象
	 * @param tbl
	 * @return
	 */
	public static TblPr getTblPr(Tbl tbl){
		TblPr tblpr = tbl.getTblPr();
		if(tblpr == null){
			tblpr = new TblPr();
			tbl.setTblPr(tblpr);
		}
		return tblpr;
	}

	/**
	 * 功能描述：取单元格配置对象
	 * @param tc 单元格对象
	 * @return
	 */
	public static TcPr getTcPr(Tc tc) {
	   TcPr tcPr = tc.getTcPr();
	   if (tcPr == null) {
	     tcPr = new TcPr();
	     tc.setTcPr(tcPr);
	   }
	   return tcPr;
	 }
	/**
	 * 功能描述：设置单元格垂直对齐方式
	 * @param tc 单元格对象
	 * @param vAlignType 单元格垂直对齐方式 STVerticalJc
	 */
	public static void setTcVAlign(Tc tc, STVerticalJc vAlignType) {
		if(vAlignType != null){
			TcPr tcPr = getTcPr(tc);
			CTVerticalJc vAlign = tcPr.getVAlign();
			if(vAlign == null){
				vAlign = new CTVerticalJc();
				tcPr.setVAlign(vAlign);
			}
			vAlign.setVal(vAlignType);
	   }
	}

	/**
	 * 功能描述：设置单元格背景色
	 * @param tc
	 * @param shdtype
	 * @param color
	 */
	public static void setTcShdStyle(Tc tc,STShd shdtype, String color ){
		if (shdtype != null) {
			TcPr tcpr = getTcPr(tc);
			CTShd shd = new CTShd();
		   shd.setVal(shdtype);
		   shd.setColor("auto");
		   shd.setFill(color);
		   tcpr.setShd(shd);
		 }
	}
	/**
	 * 功能描述：设置表格奇偶行背景色
	 * @param tbl
	 * @param shdtype
	 * @param color1
	 * @param color2
	 * @param heandersize
	 */
	public static void setOERowShdStyle(Tbl tbl, STShd shdtype, String color1, String color2, int heandersize ){
		List<Object> rowList = tbl.getContent();
		int fillindex = 1;
		for(int row=0, rowsize=rowList.size(); row<rowsize; row++,fillindex++){
			Tr tr = (Tr) XmlUtils.unwrap(rowList.get(row));
			List<Object> colList = tr.getContent();
			if(row < heandersize){
				continue;
			}

			for(int col=0, colsize=colList.size(); col<colsize; col++){
				Tc tc = (Tc) XmlUtils.unwrap(colList.get(col));

				if(fillindex%2 == 1){
					setTcShdStyle(tc,STShd.CLEAR,color1);
				}else{
					setTcShdStyle(tc,STShd.CLEAR,color2);
				}
			}
		}
	}

	/**
	 * 功能描述：Object数据转换为String类型
	 * @param obj
	 * @param defaultStr 假如obj对象为空，则返回的值
	 * @return
	 * @author cnbi
	 */
	public static String converObjToStr(Object obj , String defaultStr){
		if(obj != null){
			return obj.toString().trim();
		}
		return defaultStr.trim();
	}

	/**
	 * 创建页眉
	 * @param wordMLPackage
	 * @param content
	 * @param jcEnumeration
	 * @throws Exception
	 */
	public static void createHeader(WordprocessingMLPackage wordMLPackage, String content, JcEnumeration jcEnumeration ) throws Exception{
		Relationship relationship = createTextHeaderPart(wordMLPackage, content, jcEnumeration);
		createHeaderReference(wordMLPackage, relationship);
	}

	// 文字页面
	public static Relationship createTextHeaderPart(WordprocessingMLPackage wordMLPackage,  String content, JcEnumeration jcEnumeration) throws Exception {
		HeaderPart headerPart = new HeaderPart();
		headerPart.setJaxbElement(getTextHdr(wordMLPackage, content, jcEnumeration));
		Relationship rel = wordMLPackage.getMainDocumentPart().addTargetPart(headerPart);
		return rel;
	}

	public static Hdr getTextHdr(WordprocessingMLPackage wordMLPackage,String content, JcEnumeration jcEnumeration) throws Exception {
		Hdr hdr = new Hdr();
		hdr.getContent().add(getHeadorFooterPart(wordMLPackage, StyleConstains.Header, content, jcEnumeration));
//		hdr.getContent().add(createStyledParagraphOfText(wordMLPackage, StyleConstains.Header, ""));
		PPr pPr = ((P)hdr.getContent().get(0)).getPPr();
		Ind ind = new Ind();
		/*ind.setLeft(new BigInteger("2520"));
		ind.setFirstLine(new BigInteger("420"));
		ind.setRight(new BigInteger("440"));*/
		pPr.setInd(ind);
		Jc jc = new Jc();
		jc.setVal(JcEnumeration.RIGHT);
		pPr.setJc(jc);
		return hdr;
	}
	
	public static void createHeaderReference(WordprocessingMLPackage wordMLPackage,  Relationship relationship) {
		List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
		SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
		// There is always a section wrapper, but it might not contain a sectPr
		if (sectPr == null) {
			sectPr = new SectPr();
			wordMLPackage.getMainDocumentPart().addObject(sectPr);
			sections.get(sections.size() - 1).setSectPr(sectPr);
		}
		HeaderReference headerReference = new HeaderReference();
		headerReference.setId(relationship.getId());
		headerReference.setType(HdrFtrRef.DEFAULT);
		sectPr.getEGHdrFtrReferences().add(headerReference);
	}
	
	/**
	 * 创建页脚
	 * @param wordMLPackage
	 * @param content
	 * @param jcEnumeration
	 * @throws Exception 
	 */
	public static void createFooter(WordprocessingMLPackage wordMLPackage,  String content, JcEnumeration jcEnumeration) throws Exception{
		Relationship relationship = createTextFooterPart(wordMLPackage, content, jcEnumeration);
		createFooterReference(wordMLPackage, relationship);
	}
	
	public static Relationship createTextFooterPart(WordprocessingMLPackage wordMLPackage,  String content, JcEnumeration jcEnumeration) throws Exception {
		FooterPart footerPart = new FooterPart();
//		footerPart.setPackage(wordMLPackage);
		footerPart.setJaxbElement(getTextFtr(wordMLPackage, content, jcEnumeration));
		Relationship rel = wordMLPackage.getMainDocumentPart().addTargetPart(footerPart);
		return rel;
	}
	
	public static Ftr getTextFtr(WordprocessingMLPackage wordMLPackage,String content, JcEnumeration jcEnumeration) throws Exception {
		Ftr ftr = new Ftr();
		ftr.getContent().add(getHeadorFooterPart(wordMLPackage,StyleConstains.Footer, content, jcEnumeration));
		return ftr;
	}
	
	public static void createFooterReference(WordprocessingMLPackage wordMLPackage,  Relationship relationship){
		List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
		SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
		// There is always a section wrapper, but it might not contain a sectPr
		if (sectPr == null) {
			sectPr = new SectPr();
			wordMLPackage.getMainDocumentPart().addObject(sectPr);
			sections.get(sections.size() - 1).setSectPr(sectPr);
		}
		FooterReference footerReference = new FooterReference();
		footerReference.setId(relationship.getId());
		footerReference.setType(HdrFtrRef.DEFAULT);
		sectPr.getEGHdrFtrReferences().add(footerReference);
	}
	
	public static P getHeadorFooterPart(WordprocessingMLPackage wordprocessingMLPackage, String styleConstain , String content, JcEnumeration jcEnumeration){
		P part = createStyledParagraphOfText(wordprocessingMLPackage, styleConstain, "");
		int start = content.indexOf("${");
		int end = content.indexOf("}");
		while(start != -1){
			String cont = content.substring(0, start);
			R run = new R();
			Text text = getText(cont,null);
			run.getContent().add(text);
			part.getContent().add(run);
			cont = content.substring(start+2, end);
			addFieldSeparate(part);
			addFieldBegin(part,false);
			if("PAGE".equals(cont) || "NUMPAGES".equals(cont)){
				cont += " \\* MERGEFORMAT ";
			}
			addFieldContent(part, cont);//"PAGE  \\* MERGEFORMAT "
			addFieldSeparate(part);
			addFieldEnd(part);
			content = content.substring(end+1);
			start = content.indexOf("${");
			end = content.indexOf("}");
		}
		R run = new R();
		Text text = getText(content,null);
		run.getContent().add(text);
		part.getContent().add(run);
		setParaJcAlign(part,jcEnumeration);
		return part;
	}

	/**
	 * 设置默认样式
	 * @param wordPackage
	 * @throws JAXBException
	 */
	public static void setDefStyles(WordprocessingMLPackage wordPackage, String stylexml) throws JAXBException {
		
		java.io.InputStream is = null;
		try {
			is = new FileInputStream(new File(stylexml));
			StyleDefinitionsPart styledf = wordPackage.getMainDocumentPart().getStyleDefinitionsPart();
			Styles styles =  styledf.unmarshal( is );
			styledf.setContents(styles);
			
			PropertyResolver pr = wordPackage.getMainDocumentPart().getPropertyResolver();
			List<Style> styleslist = styles.getStyle();
			for(int i = 0 , size = styleslist.size(); i < size; i++){
//				logger(i+"==>");
				Style style = styleslist.get(i);
//				logger(style.getStyleId());
//				logger("==>");
//				logger(style.getName().getVal()+"\n");
				pr.activateStyle(style);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} catch (JAXBException e) {
			throw e;
		}
		
	}
	
	/**
	 * 设置编号
	 * @param wordMLPackage
	 * @param numnberxml
	 * @throws JAXBException
	 */
	public static void setDefNumber( WordprocessingMLPackage wordMLPackage, String numnberxml) throws JAXBException {
		java.io.InputStream is = null;
		try {
			NumberingDefinitionsPart numbPart = wordMLPackage.getMainDocumentPart().getNumberingDefinitionsPart();
			is = new FileInputStream(new File(numnberxml));
			Numbering numb =  numbPart.unmarshal(is);
			numbPart.setContents(numb);
//			System.out.println("numbPart");
		} catch (IOException e) {
			e.printStackTrace();
		} catch (JAXBException e) {
			throw e;
//		} finally {
//			if(is != null){
//				try {
//					is.close();
//				} catch (IOException e) {
//				}
//				is = null;
//			}
		}
	}
	
	public static void logger(String m){
		System.out.print(m);
	}
	
	
}
