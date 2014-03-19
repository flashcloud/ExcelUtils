/*
 * Copyright 2003-2005 try2it.com.
 * Created on 2005-7-7
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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import net.sf.excelutils.ExcelParser;
import net.sf.excelutils.WorkbookUtils;

import org.apache.commons.beanutils.DynaBean;
import org.apache.commons.beanutils.DynaProperty;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import bsh.EvalError;
import bsh.Interpreter;

/**
 * <p>
 * <b>IfTag</b> is a class which parse #if #else #end
 * </p>
 * 
 * @author rainsoft
 * @version $Revision: 123 $ $Date: 2006-09-17 15:03:24 +0800 (星期日, 17 九月 2006) $
 */
@SuppressWarnings("unchecked")
public class IfTag implements ITag {
	public static final String KEY_IF = "#if";

	public static final String KEY_END = "#end";

	private Log LOG = LogFactory.getLog(IfTag.class);

	public int[] parseTag(Object context, HSSFWorkbook wb, HSSFSheet sheet, HSSFRow curRow, HSSFCell curCell) {
		int ifstart = curRow.getRowNum();
		int ifend = -1;
		int ifCount = 0;
		String ifstr = "";
		boolean bFind = false;
		for (int rownum = ifstart; rownum <= sheet.getLastRowNum(); rownum++) {
			HSSFRow row = sheet.getRow(rownum);
			if (null == row)
				continue;
			for (short colnum = row.getFirstCellNum(); colnum <= row.getLastCellNum(); colnum++) {
				HSSFCell cell = row.getCell(colnum, HSSFRow.RETURN_NULL_AND_BLANK);
				if (null == cell)
					continue;
				if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
					String cellstr = cell.getStringCellValue();

					// get the tag instance for the cellstr
					ITag tag = ExcelParser.getTagClass(cellstr);

					if (null != tag) {
						if (tag.hasEndTag()) {
							if (0 == ifCount) {
								ifstart = rownum;
								ifstr = cellstr;
							}
							ifCount++;
							break;
						}
					}
					if (cellstr.startsWith(KEY_END)) {
						ifend = rownum;
						ifCount--;
						if (ifstart >= 0 && ifend >= 0 && ifend > ifstart && ifCount == 0) {
							bFind = true;
						}
						break;
					}
				}
			}
			if (bFind)
				break;
		}

		if (!bFind)
			return new int[] { 0, 0, 1 };

		// test if condition
		boolean bResult = false;
		// remove #if tag and get condition expression
		String expr = ifstr.trim().substring(KEY_IF.length()).trim();

		// parse the condition expression
		expr = (String) ExcelParser.parseStr(context, expr, true);

		// use beanshell to eval expression value

		try {
			Interpreter in = createInterpreter(context);
			LOG.debug("IfTag test expr=" + expr);
			Object v = in.eval(expr);
			bResult = ((Boolean) v).booleanValue();
		} catch (Exception e) {
			LOG.error("IfTag test expr error", e);
			bResult = false;
		}

		if (bResult) { // if condition is true
			// remove #if tag and #end tag only
			sheet.removeRow(WorkbookUtils.getRow(ifstart, sheet));
			sheet.removeRow(WorkbookUtils.getRow(ifend, sheet));
			// remove merged region in ifstart & ifend
			for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
				CellRangeAddress r = sheet.getMergedRegion(i);
				if (r.getFirstRow() == ifstart && r.getLastRow() == ifstart || r.getFirstRow() == ifend
						&& r.getLastRow() == ifend) {
					sheet.removeMergedRegion(i);
					// we have to back up now since we removed one
					i = i - 1;
				}
			}
			if (ifend + 1 <= sheet.getLastRowNum()) {
				sheet.shiftRows(ifend + 1, sheet.getLastRowNum(), -1, true, true);
			}
			if (ifstart + 1 <= sheet.getLastRowNum()) {
				sheet.shiftRows(ifstart + 1, sheet.getLastRowNum(), -1, true, true);
			}
			return new int[] { 1, -2, 1 };
		} else { // if condition is false
			// remove #if #end block
			for (int rownum = ifstart; rownum <= ifend; rownum++) {
				sheet.removeRow(WorkbookUtils.getRow(rownum, sheet));
			}
			// remove merged region in ifstart & ifend
			for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
				CellRangeAddress r = sheet.getMergedRegion(i);
				if (r.getFirstRow() >= ifstart && r.getLastRow() <= ifend) {
					sheet.removeMergedRegion(i);
					// we have to back up now since we removed one
					i = i - 1;
				}
			}
			if (ifend + 1 <= sheet.getLastRowNum()) {
				sheet.shiftRows(ifend + 1, sheet.getLastRowNum(), -(ifend - ifstart + 1), true, true);
			}
			return new int[] { ExcelParser.getSkipNum(ifstart, ifend), ExcelParser.getShiftNum(ifend, ifstart), 1 };
		}
	}

	public String getTagName() {
		return KEY_IF;
	}

	public boolean hasEndTag() {
		return true;
	}

	/**
	 * create interpreter
	 * 
	 * @param context
	 * @return
	 * @throws EvalError
	 */
	private Interpreter createInterpreter(Object context) throws EvalError {
		Interpreter in = new Interpreter();
		// insert object into bsh
		if (context instanceof Map) {
			for (Iterator it = ((Map) context).entrySet().iterator(); it.hasNext();) {
				Entry entry = (Entry) it.next();
				in.set(entry.getKey().toString(), entry.getValue());
			}
		} else if (context instanceof DynaBean) {
			DynaProperty properties[] = ((DynaBean) context).getDynaClass().getDynaProperties();
			for (int i = 0; i < properties.length; i++) {
				Object value = ((DynaBean) context).get(properties[i].getName());
				in.set(properties[i].getName(), value);
			}
		}
		return in;
	}

	public static void main(String[] args) {
		Interpreter in = new Interpreter();
		try {
			in.set("a", "5a");
			in.set("b", "5a");
			List errors = new ArrayList();
			errors.add("abc");
			Map context = new HashMap();
			context.put("errors", errors);

			in.set("context", context);

			in.eval("bar=\"5a\"==\"5a\"");
			System.out.println(in.get("bar"));
			System.out.println("ab${addd}dd${cccc}aa".replaceAll("\\$\\{|\\}", ""));
			System.out.println(in.eval("(context.get(\"errors\").size() == 2)"));
			System.out.println(in.eval("abc == null"));
		} catch (EvalError e) {
			e.printStackTrace();
		}
	}
}
