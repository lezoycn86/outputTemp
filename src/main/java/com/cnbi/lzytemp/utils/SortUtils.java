package com.cnbi.lzytemp.utils;

import java.util.List;

/**
 * @Title: SortUtils
 * @Description: 排序
 * @author: lzy
 * @date: 2019/9/5 11:50
 */
public class SortUtils {
    public static List<String> listSort(List<String> input) {
        for (int i = 0; i < input.size() - 1; i++) {
            for (int j = 0; j < input.size() - i - 1; j++) {
                if (isBiggerThan(input.get(j), input.get(j + 1))) {
                    String temp = input.get(j);
                    input.set(j,input.get(j + 1));
                    input.set(j + 1, temp);
                }
            }
        }
        return input;
    }
    private static boolean isBiggerThan(String first, String second){
        if(first==null||second==null||"".equals(first) || "".equals(second)){
            System.out.println("字符串不能为空！");
            return false;
        }
        char[] arrayfirst=first.toCharArray();
        char[] arraysecond=second.toCharArray();
        int minSize = Math.min(arrayfirst.length, arraysecond.length);
        for (int i=0;i<minSize;i++) {
            if((int)arrayfirst[i]>(int)arraysecond[i]){
                return true;
            }else if((int)arrayfirst[i] < (int)arraysecond[i]){
                return false;
            }
        }
        if(arrayfirst.length>arraysecond.length){
            return true;
        }else {
            return false;
        }
    }
}
