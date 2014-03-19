/*
 * FileName：ExcelIterator.java 2009-6-17 
 * Copyright (C) 2003-2007 try2it.com
 */
package com.ky.ds.utils;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

/**
 * <p>
 * <b>ExcelIterator</b>is Excel的迭代器
 * </p>
 * 
 * @author rainsoft
 * @since 2009-6-17
 * @version $Revision$ $Date: 2009-06-29 09:13:42 +0800 (星期一, 2009-06-29) $
 */
public class ExcelIterator implements Iterator {
	protected HSSFSheet sheet;

	protected int index = 0;

	public ExcelIterator(HSSFSheet sheet) {
		this.sheet = sheet;
		index = sheet.getFirstRowNum();
	}

	protected Object get(int idx) {
		HSSFRow row = sheet.getRow(idx);
		List datas = new ArrayList();
		if (null != row) {
			for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
				HSSFCell cell = row.getCell((short) i);
				Object value = null;
				if (null != cell) {
					value = getCellValue(cell);
				}
				datas.add(value);
			}
		}
		return datas;
	}

	public Object next() {
		return get(index++);
	}

	public boolean hasNext() {
		if (index > sheet.getLastRowNum()) {
			return false;
		}
		return true;
	}

	public void remove() {
		throw new IllegalAccessError("remove method is not supported");
	}

	/**
	 * 获取表达式的值
	 * 
	 * @param cell 单元格
	 * @return Object
	 */
	private Object getCellValue(HSSFCell cell) {
		Object value = null;
		if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
			value = cell.getStringCellValue();
			if (null != value) {
				value = value.toString().trim();
			}
		}
		if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				value = cell.getDateCellValue();
			} else {
				value = new Double(cell.getNumericCellValue());
			}
		}
		if (cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
			value = new Boolean(cell.getBooleanCellValue());
		}
		if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
			value = String.valueOf(cell.getNumericCellValue());
			if (value == null || value == "" || "NaN".equals(value)) {
				value = cell.getStringCellValue();
			} else {
				value = new Double(cell.getNumericCellValue());
			}
		}
		if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
			value = "";
		}
		return value;
	}
}
