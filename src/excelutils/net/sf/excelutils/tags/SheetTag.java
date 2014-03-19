/*
 * Copyright 2003-2005 ExcelUtils http://excelutils.sourceforge.net
 * Created on 2005-6-22
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License. 
 */

package net.sf.excelutils.tags;

import java.util.Iterator;
import java.util.StringTokenizer;

import net.sf.excelutils.ExcelException;
import net.sf.excelutils.ExcelParser;
import net.sf.excelutils.ExcelUtils;
import net.sf.excelutils.WorkbookUtils;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * <p>
 * <b>SheetTag</b>is a class whick parse #sheet tag
 * </p>
 * 
 * @author rainsoft
 * @version $Revision: 127 $ $Date: 2006-10-13 23:48:46 +0800 (星期五, 13 十月 2006) $
 */
@SuppressWarnings("unchecked")
public class SheetTag implements ITag {

	public static final String KEY_SHEET = "#sheet";

	private Log LOG = LogFactory.getLog(SheetTag.class);

	/**
	 * Parse #sheet detail in list by sheetName
	 */
	public int[] parseTag(Object context, Workbook wb, Sheet sheet, Row curRow, Cell curCell) throws ExcelException {
		String sheetExpr = curCell.getStringCellValue();
		StringTokenizer st = new StringTokenizer(sheetExpr, " ");

		String properties = "";
		String property = "";
		String sheetName = "";
		// parse the collection an object
		int pos = 0;
		while (st.hasMoreTokens()) {
			String str = st.nextToken();
			if (pos == 1) {
				property = str;
			}
			if (pos == 3) {
				properties = str;
			}
			if (pos == 5) {
				sheetName = str;
			}
			pos++;
		}

		// get collection
		Object collection = ExcelParser.parseStr(context, properties);
		if (null == collection) {
			return new int[] { 0, 0, 1 };
		}

		// remove #sheet tag
		sheet.removeRow(curRow);

		// remove merged region in forstart & forend
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			CellRangeAddress r = sheet.getMergedRegion(i);
			if (r.getFirstRow() >= curRow.getRowNum() && r.getLastRow() <= curRow.getRowNum()) {
				sheet.removeMergedRegion(i);
				i = i - 1;
			}
		}
		sheet.shiftRows(curRow.getRowNum() + 1, sheet.getLastRowNum(), -1, true, true);

		// get the iterator of collection
		Iterator iterator = ExcelParser.getIterator(collection);
		if (null != iterator) {
			// first obj, use parse method
			Object firstObj = null;
			if (iterator.hasNext()) {
				firstObj = iterator.next();
			}

			// next obj, clone sheet and use parseSheet method
			while (iterator.hasNext()) {
				Object obj = iterator.next();
				ExcelUtils.addValue(context, property, obj);
				try {
					int sheetIndex = WorkbookUtils.getSheetIndex(wb, sheet);

					// clone sheet
					Sheet cloneSheet = wb.cloneSheet(sheetIndex);

					// set cloneSheet name
					int cloneSheetIndex = WorkbookUtils.getSheetIndex(wb, cloneSheet);
					setSheetName(obj, wb, cloneSheetIndex, sheetName);

					// parse cloneSheet
					ExcelUtils.parseSheet(context, wb, cloneSheet);
				} catch (Exception e) {
					if (LOG.isErrorEnabled()) {
						LOG.error("parse sheet error", e);
					}
				}
			}

			if (null != firstObj) {
				ExcelUtils.addValue(context, property, firstObj);
				// set sheet name
				int sheetIndex = WorkbookUtils.getSheetIndex(wb, sheet);
				setSheetName(firstObj, wb, sheetIndex, sheetName);
			}
		}

		return new int[] { 0, -1, 0 };
	}

	private void setSheetName(Object context, Workbook wb, int sheetIndex, String sheetName) {
		// set sheetName
		if (!"".equals(sheetName)) {
			Object o = ExcelParser.getValue(context, sheetName);
			if (null != o && !"".equals(o)) {
				sheetName = o.toString();
			}
			wb.setSheetName(sheetIndex, sheetName);
		}
	}

	public boolean hasEndTag() {
		return false;
	}

	public String getTagName() {
		return KEY_SHEET;
	}
}
