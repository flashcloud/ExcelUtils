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
package net.sf.excelutils.webwork;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Iterator;

import javax.servlet.ServletContext;

import net.sf.excelutils.ExcelUtils;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import com.opensymphony.webwork.config.Configuration;
import com.opensymphony.xwork.ActionSupport;
import com.opensymphony.xwork.ObjectFactory;
import com.opensymphony.xwork.util.OgnlValueStack;

/**
 * @author <a href="mailto:joke_way@yahoo.com.cn">jokeway</a>
 * @since 2005-10-10
 * @version $Revision: 126 $ $Date: 2006-09-17 16:32:41 +0800 (星期日, 17 九月 2006) $
 */
@SuppressWarnings("unchecked")
public class ExcelManager {
	private static final Log log = LogFactory.getLog(ExcelManager.class);

	private static ExcelManager instance = null;

	protected ExcelLoader excelLoader = null;

	public final static synchronized ExcelManager getInstance() {
		if (instance == null) {
			String classname = ExcelManager.class.getName();

			if (Configuration.isSet("webwork.excel.manager.classname")) {
				classname = Configuration.getString("webwork.excel.manager.classname").trim();
			}

			try {
				log.info("Instantiating Excel ConfigManager!, " + classname);
				instance = (ExcelManager) ObjectFactory.getObjectFactory().buildBean(Class.forName(classname), null);
			} catch (Exception e) {
				log.fatal("Fatal exception occurred while trying to instantiate a Excel ConfigManager instance, " + classname,
						e);
			}
		}

		// if the instance creation failed, make sure there is a default instance
		if (instance == null) {
			instance = new ExcelManager();
		}

		return instance;
	}

	public Object buildContextObject(OgnlValueStack stack) {
		// add action properties to default context
		if (null != stack) {
			for (Iterator it = stack.getRoot().iterator(); it.hasNext();) {
				Object obj = it.next();
				if (null != obj && obj instanceof ActionSupport) {
					Field[] field = obj.getClass().getDeclaredFields();
					for (int i = 0; null != field && i < field.length; i++) {
						Method method = null;
						try {
							method = obj.getClass().getMethod(
									"get" + field[i].getName().substring(0, 1).toUpperCase() + field[i].getName().substring(1),
									new Class[] {});
						} catch (Exception e) {
							method = null;
						}
						if (null == method) {
							try {
								method = obj.getClass().getMethod(
										"is" + field[i].getName().substring(0, 1).toUpperCase() + field[i].getName().substring(1),
										new Class[] {});
							} catch (Exception e) {
								method = null;
							}
						}
						if (null != method) {
							Object value = stack.findValue(field[i].getName());
							ExcelUtils.getContext().set(field[i].getName(), value);
						}
					}
				}
			}
		}
		return ExcelUtils.getContext();
	}

	protected ExcelLoader getExcelLoader(ServletContext context) {
		if (excelLoader == null) {
			excelLoader = new MutiExcelLoader(new ExcelLoader[] { new WebappExcelLoader(context) });
		}
		return excelLoader;
	}

	public InputStream getExcel(ServletContext context, String location) {
		return getExcelLoader(context).getExcel(location);
	}

}
