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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * <p>
 * <b>ITag </b> is a interface which define the tag
 * </p>
 * 
 * @author rainsoft
 * @version $Revision: 112 $ $Date: 2006-08-22 18:54:05 +0800 (星期二, 22 八月 2006) $
 */
public interface ITag {

  /**
   * parse the tag
   * 
   * @param context data object
   * @param wb excel workbook
   * @param sheet excel sheet
   * @param curRow excel row
   * @param curCell excel cell
   * @return int[] {skip number, shift number, break flag}
   */
  public int[] parseTag(Object context, Workbook wb, Sheet sheet, Row curRow, Cell curCell) throws ExcelException;

  /**
   * tag has #end flag
   * 
   * @return boolean
   */
  public boolean hasEndTag();

  /**
   * get the tag name
   * 
   * @return str
   */
  public String getTagName();
}
