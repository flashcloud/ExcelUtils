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

import net.sf.excelutils.ExcelException;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * <p>
 * <b>PageTag </b> is a class which parse the #page tag Because a bug in POI, you must place a split char in your sheet
 * near the #page tag
 * </p>
 * 
 * @author rainsoft
 * @version $Revision: 127 $ $Date: 2006-10-13 23:48:46 +0800 (星期五, 13 十月 2006) $
 */
public class PageTag implements ITag {

	private Log LOG = LogFactory.getLog(IfTag.class);

	public static final String KEY_PAGE = "#page";

	public int[] parseTag(Object context, Workbook wb, Sheet sheet, Row curRow, Cell curCell) throws ExcelException {
		int rowNum = curRow.getRowNum();
		LOG.debug("#page at rownum = " + rowNum);
		sheet.setRowBreak(rowNum - 1);
		sheet.removeRow(curRow);
		if (rowNum + 1 <= sheet.getLastRowNum()) {
			sheet.shiftRows(rowNum + 1, sheet.getLastRowNum(), -1, true, true);
		}
		return new int[] { 0, -1, 0 };
	}

	public String getTagName() {
		return KEY_PAGE;
	}

	public boolean hasEndTag() {
		return false;
	}
}
