/*
 * Copyright 2003-2005 ExcelUtils http://excelutils.sourceforge.net
 * Created on 2005-6-18
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

package net.sf.excelutils;

/**
 * <p>
 * <b>ExcelException </b>
 * </p>
 * 
 * @author rainsoft
 * @version $Revision: 112 $ $Date: 2006-08-22 18:54:05 +0800 (星期二, 22 八月 2006) $
 */
public class ExcelException extends Exception {
  
  private static final long serialVersionUID = 1L;

  public ExcelException() {
    super();
  }

  /**
   * ExcelException constructor
   * 
   * @param s exception clause
   */
  public ExcelException(String s) {
    super(s);
  }
  
  public ExcelException(String s, Exception e) {
  	super(s, e);
  }
}
