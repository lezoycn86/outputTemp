package com.cnbi.lzytemp.utils;

public class Docx4jToolUtils {
	
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
	
	public static boolean isNumb(String numb){
		boolean ret = true;
		numb = numb.replaceAll(",", "").toUpperCase();
		if(numb.contains("E")){
			String[] numbs =	numb.split("E");
			if(numbs.length < 3){
				for(String s : numbs){
						ret &= isNumb(s);
				}
			}else{
				ret =false;
			}
		}else{
			ret = numb.matches("(-|\\+|)[0-9]+(\\.|)[0-9]*");
			if(isStringNumb(numb)){
				ret = false;
			}
		}
		return ret;
	}
	
	public static boolean isStringNumb(String numb){
		boolean ret = true;
		numb = numb.replaceAll(",", "").toUpperCase();
		ret = numb.matches("(-|\\+|)0[^\\.].*");
		return ret ;
	}
	
}
