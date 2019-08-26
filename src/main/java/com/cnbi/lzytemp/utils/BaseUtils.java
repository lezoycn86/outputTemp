package com.cnbi.lzytemp.utils;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.OpcPackage;

import java.io.File;
import java.io.InputStream;

public class BaseUtils {
	
	/** 
	 * 功能描述：加载文档信息 
	 * @param file     文件
	 * @param password 密码
	 * @throws Docx4JException 
	 * @author cnbi 
	 */
	public static OpcPackage loadOpcPackage(File file, String password) throws Docx4JException {
		return OpcPackage.load(file, password);
	}
	
	/**
	 * 功能描述：加载文档信息 
	 * @param input 文件
	 * @param password 密码
	 * @return
	 * @throws Docx4JException
	 */
	public static OpcPackage loadOpcPackage(InputStream input, String password) throws Docx4JException {
		return OpcPackage.load(input, password);
	}
	
	
	/** 
	 * 功能描述：保存文档信息 
	 * @param oPackage  文档处理包对象 
	 * @param file         文件 
	 * @throws Docx4JException 
	 * @author cnbi 
	 */
	public static void savePackage(OpcPackage opackage, File file) throws Docx4JException {
		opackage.save(file);
	}
	
}
