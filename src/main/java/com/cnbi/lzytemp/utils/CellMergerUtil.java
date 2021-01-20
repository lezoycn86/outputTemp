package com.cnbi.lzytemp.utils;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;
import java.util.Map.Entry;

public class CellMergerUtil {

	/**
	 * <pre>
	 * 根据集合数据和需要合并的列索引号返回所有列的所有待合并区域
	 * </pre>
	 */
	public static Map<Integer, List<CellRangeAddress>> getRanglesByData(
			List<Map<String, Object>> itemList, int[] mergerRows, int startRows) {
		Map<Integer, List<CellRangeAddress>> ranglesMap = new LinkedHashMap<Integer, List<CellRangeAddress>>();
		if(mergerRows == null || mergerRows.length == 0){
			return ranglesMap;
		}
		if(CollectionUtils.isEmpty(itemList)){
			return ranglesMap;
		}
		//根据数据索引号获取数据的合并位置
		for(int i=0 , lens=mergerRows.length ; i < lens ; i++){
			int size = itemList.size();
			for(int j=0 ; j < size ; j++){
				int merger = 0;
				Map<String , Object> item = itemList.get(j);
				String value = getValuesByMergerIndex(item , mergerRows , i);
				for(int k=j+1 ; k < size ; k++){
					Map<String , Object> nextItem = itemList.get(k);
					String nextValue = getValuesByMergerIndex(nextItem , mergerRows , i);
					if(!value.equals(nextValue)){
						break;
					}
					merger++;//累计需要合并单元格的行数
				}
				//使用索引从0开始的合并
				if(merger >= 1){
					//添加需要处理单元格的位置
					int column = mergerRows[i];
					int firstRow = startRows + j;
					int lastRow = firstRow + merger;
					CellRangeAddress range = new CellRangeAddress(firstRow, lastRow, column, column);
					List<CellRangeAddress> list = ranglesMap.get(column);
					if(list == null){
						list = new ArrayList<CellRangeAddress>();
					}
					list.add(range);
					ranglesMap.put(column, list);
					j = j + merger;//将待合并的区域数据移至下面合并完后的一行
				}
			}
		}
		return ranglesMap;
	}
	
	//根据值转换为对应索引列的值
	public static String getValuesByMergerIndex(Map<String , Object> item , int[] mergerRows , int i){
		StringBuilder builder = new StringBuilder();
		Object values[] = item.entrySet().toArray();
		for(int t=0 ; t < mergerRows.length ; t++){
			if(t <= i){
				builder.append(values[mergerRows[t]]).append("");
			}
		}
		return builder.toString();
	}
	

	/**
	 * <pre>
	 * 合并单元格
	 * </pre>
	 */
	public static void mergerColumns(Sheet sheet, Map<Integer, List<CellRangeAddress>> rangeMap) {
		Set<Entry<Integer , List<CellRangeAddress>>> entrySet = rangeMap.entrySet();
		for (Entry<Integer, List<CellRangeAddress>> entry : entrySet) {
			//合并单元格
			List<CellRangeAddress> valueList = entry.getValue();
			for (CellRangeAddress range : valueList) {
				sheet.addMergedRegion(range);
			}
		}
	}

}
