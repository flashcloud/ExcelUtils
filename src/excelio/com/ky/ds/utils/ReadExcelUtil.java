/*
 * FileName：ReadExcelUtil.java 2009-6-17 
 * Copyright (C) 2003-2007 try2it.com
 */
package com.ky.ds.utils;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.sql.Connection;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import net.sf.excelutils.WorkbookUtils;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.input.SAXBuilder;
import org.jdom.xpath.XPath;

import com.ky.core.exception.CoreException;
import com.ky.core.util.DateFormatUtil;
import com.ky.ds.bo.excelMapping.VendorQuoteMapping;
import com.ky.fuse.report.query.util.Database;
import com.ky.fuse.report.query.util.DesignQuery;
import com.ky.fuse.report.query.util.RowSetDynaClass;

/**
 * <p>
 * <b>ReadExcelUtil</b>is 读Excel的类
 * </p>
 * 
 * @author rainsoft
 * @since 2009-6-17
 * @version $Revision$ $Date: 2009-09-19 13:46:52 +0800 (星期六, 2009-09-19) $
 */
public class ReadExcelUtil {

	private static Log log = LogFactory.getLog(ReadExcelUtil.class);

	/**
	 * 从Excel读取数据插入数据表
	 * 
	 * @param excelIn Excel流
	 * @param sheetIndex XXXXXXXXXXXXXXXXXXXXXX
	 * @param xmlConfig XML配置文件路径
	 * @return
	 */
	public static String[] saveOrUpdateTable(Connection conn, InputStream excelIn, int sheetIndex, String xmlConfig) {
		try {
			conn.setAutoCommit(false);
			Database db = new Database(conn);

			// 获取excel的iterator
			Iterator[] its = ReadExcelUtil.getIterator(excelIn, new int[] { sheetIndex });

			List tableList = new ArrayList();
			// 获取XML配置和excel列头
			Object[] tables = getTable(its[0], xmlConfig);
			Element detailTable = (Element) tables[0];
			String[] columns = (String[]) tables[1];
			Element root = (Element) tables[2];
			// 看看是否有关联的主表
			String master = detailTable.getAttributeValue("master");
			if (null != master && !"".equals(master)) {
				Element masterTable = (Element) XPath.selectSingleNode(root, "/config/tables/table[@name='" + master + "']");
				if (null == masterTable) {
					throw new CoreException("找不到XML中master属性指定的" + master + "配置", null);
				}
				tableList.add(masterTable);
			}
			tableList.add(detailTable); // detail

			String globalRepeatLastNo = "";

			// 主从表的映射NO
			Map masterNoMap = new HashMap();
			String[] result = new String[tableList.size() + 1];
			String[] insertSql = new String[tableList.size()];
			String[] updateSql = new String[tableList.size()];
			List[] xmlColumns = new List[tableList.size()];
			List[] preSqls = new List[tableList.size()];
			List[] postSqls = new List[tableList.size()];
			List[] validSqls = new List[tableList.size()];
			int[][] mapping = new int[tableList.size()][];
			int[] insert_records = new int[tableList.size()];
			int[] update_records = new int[tableList.size()];
			// 使用该Map用来判断数据是否是重复了，自动忽略重复的数据
			Map[] keyUniqueMap = new Map[tableList.size()];

			for (int tableIndex = 0; tableIndex < tableList.size(); tableIndex++) {
				Element table = (Element) tableList.get(tableIndex);
				// 根据XML配置组织SQL语句
				String tableName = table.getAttributeValue("table");
				String title = table.getAttributeValue("nick");
				insertSql[tableIndex] = "INSERT INTO " + tableName + "(";
				String insertValue = " VALUES(";
				updateSql[tableIndex] = "UPDATE " + tableName + " SET ";
				String whereSql = " WHERE ";

				List preSql = XPath.selectNodes(table, "sqls/pre-sql");
				for (int aa = 0; preSql != null && aa < preSql.size(); aa++) {
					String sql = ((Element) preSql.get(aa)).getTextNormalize();
					if (StringUtils.isNotEmpty(sql)) {
						if (preSqls[tableIndex] == null) {
							preSqls[tableIndex] = new ArrayList();
						}
						preSqls[tableIndex].add(sql);
					}
				}

				List postSql = XPath.selectNodes(table, "sqls/post-sql");
				for (int aa = 0; postSql != null && aa < postSql.size(); aa++) {
					String sql = ((Element) postSql.get(aa)).getTextNormalize();
					if (StringUtils.isNotEmpty(sql)) {
						if (postSqls[tableIndex] == null) {
							postSqls[tableIndex] = new ArrayList();
						}
						postSqls[tableIndex].add(sql);
					}
				}

				List validSql = XPath.selectNodes(table, "sqls/valid-sql");
				for (int aa = 0; validSql != null && aa < validSql.size(); aa++) {
					String sql = ((Element) validSql.get(aa)).getTextNormalize();
					if (StringUtils.isNotEmpty(sql)) {
						if (validSqls[tableIndex] == null) {
							validSqls[tableIndex] = new ArrayList();
						}
						validSqls[tableIndex].add(sql);
					}
				}

				xmlColumns[tableIndex] = XPath.selectNodes(table, "columns/column");
				mapping[tableIndex] = new int[xmlColumns[tableIndex].size()];
				for (int i = 0; i < xmlColumns[tableIndex].size(); i++) {
					Element column = (Element) xmlColumns[tableIndex].get(i);
					String columnName = column.getAttributeValue("name");
					String nickName = column.getAttributeValue("nick");

					// 寻找XML中的定义与Excel中的映射关系
					mapping[tableIndex][i] = -1;
					// 根据NICKNAME寻找数据
					for (int j = 0; j < columns.length; j++) {
						if (columns[j].equals(nickName)) {
							mapping[tableIndex][i] = j;
							break;
						}
					}

					// 先组织SQL
					insertSql[tableIndex] += columnName + ",";
					if (!"true".equals(column.getAttributeValue("primary"))) {
						updateSql[tableIndex] += columnName + "=$P{" + columnName + "},";
					} else {
						whereSql += columnName + " =$P{" + columnName + "} AND ";
					}
					insertValue += "$P{" + columnName + "},";
				}
				insertSql[tableIndex] = insertSql[tableIndex].substring(0, insertSql[tableIndex].length() - 1);
				insertValue = insertValue.substring(0, insertValue.length() - 1);
				insertSql[tableIndex] += ")" + insertValue + ")";
				updateSql[tableIndex] = updateSql[tableIndex].substring(0, updateSql[tableIndex].length() - 1);
				whereSql = whereSql.substring(0, whereSql.length() - 4);
				updateSql[tableIndex] += whereSql;
			}

			int lineNo = 2;
			while (its[0].hasNext()) {
				List datas = (List) its[0].next();
				// 判断datas是否都是空白行
				String line = "";
				for (int i = 0; i < datas.size(); i++) {
					Object obj = datas.get(i);
					if (obj != null) {
						line += obj.toString();
					}
				}
				if (null == line || "".equals(line)) {
					continue;
				}

				if (datas.size() > 0) {
					lineNo++;

					for (int tableIndex = 0; tableIndex < tableList.size(); tableIndex++) {

						String isql = insertSql[tableIndex];
						String usql = updateSql[tableIndex];
						boolean insert = false;
						List keyParams = new ArrayList();

						Map rowData = new HashMap();
						for (int i = 0; i < xmlColumns[tableIndex].size(); i++) {

							Object value = null;
							// 读取excel的数据
							if (mapping[tableIndex][i] >= 0) {

								value = datas.get(mapping[tableIndex][i]);

							}

							Element column = (Element) xmlColumns[tableIndex].get(i);
							String columnName = column.getAttributeValue("name");
							String primary = column.getAttributeValue("primary");
							String nickName = column.getAttributeValue("nick");
							String def = column.getAttributeValue("default");
							String not_null = column.getAttributeValue("not-null");
							String length = column.getAttributeValue("length");
							String max = column.getAttributeValue("max");
							String min = column.getAttributeValue("min");
							String type = column.getAttributeValue("type");
							String enumCheck = column.getAttributeValue("enumCheck");
							String updateFlag = column.getAttributeValue("update");
							String insertOnly = column.getAttributeValue("insert-only");
							// 是否复制上一行的值
							String repeatLast = column.getAttributeValue("repeat-last");

							// 校验数据
							if (StringUtils.isNotEmpty(not_null) && "true".equals(not_null)) {
								if ((null == value || "".equals(value)) && StringUtils.isEmpty(def)) {
									conn.rollback();
									return new String[] { "第" + lineNo + "行，'" + nickName + "'的值不能为空" };
								}
							}

							if (StringUtils.isNotEmpty(type)) {
								// 20090806 ljy long 型字段如果没有数据，传过来是"" ??
								if ("long".equals(type) && value instanceof String && "".equals(value)) {
									value = null;
								}
								if ("int".equals(type) && value instanceof String && "".equals(value)) {
									value = null;
								}
								if ("number".equals(type) && value instanceof String && "".equals(value)) {
									value = null;
								}
								if ("date".equals(type) && value instanceof String && "".equals(value)) {
									value = null;
								}
								if (null != value) {
									// 20090807 ljy
									// 手工输的"true"或"false"导进来无法转成Boolean
									if ("bool".equals(type)
											&& value instanceof String
											&& ("TRUE".equals(((String) value).toUpperCase()) || "FALSE".equals(((String) value)
													.toUpperCase()))) {
										value = new Boolean("TRUE".equals(((String) value).toUpperCase()));
									}
									// 20090806 ljy number只能转成double
									if ("long".equals(type) && value instanceof Number) {
										value = new Long(((Number) value).longValue());
									}
									if ("int".equals(type) && value instanceof Number) {
										value = new Integer(((Number) value).intValue());
									}
									if ("string".equals(type) && !(value instanceof String) || "number".equals(type)
											&& !(value instanceof Number) || "date".equals(type) && !(value instanceof Date)
											|| "int".equals(type) && !(value instanceof Integer) || "long".equals(type)
											&& !(value instanceof Long) || "bool".equals(type) && !(value instanceof Boolean)) {

										boolean ok = false;
										// 日期格式可以是yyyy-MM-dd或yyyyMMdd
										if ("date".equals(type) && value instanceof String) {
											try {
												value = DateFormatUtil.parseDateTime(value.toString());
												ok = true;
											} catch (Throwable e) {
												try {
													value = DateFormatUtil.parse(value.toString(), "yyyyMMdd");
													ok = true;
												} catch (Throwable ex) {
												}
											}
										}
										// 如果是数字，则日期格式是yyyyMMdd或者是毫秒数
										if ("date".equals(type) && value instanceof Number) {
											try {
												// double to string 会有科学计数法
												value = DateFormatUtil.parse(Long.toString(((Number) value).longValue()), "yyyyMMdd");
												ok = true;
											} catch (Throwable ex) {
												try {
													value = new Date(((Number) value).longValue());
													ok = true;
												} catch (Throwable e) {
												}
											}
										}
										// 如果要求字符串，但是值是数字型的，则转为字符串
										if ("string".equals(type) && value instanceof Number) {
											value = Long.toString(((Number) value).longValue());
											ok = true;
										}

										// 如果要求的是数字，但是值是字符串，则转为数字
										if ("number".equals(type) && value instanceof String) {
											try {
												value = new Double(Double.parseDouble(value.toString()));
												ok = true;
											} catch (Exception ex) {
											}
										}
										if ("int".equals(type) && value instanceof String) {
											try {
												value = new Integer(Integer.parseInt(value.toString()));
												ok = true;
											} catch (Exception ex) {
											}
										}
										if ("long".equals(type) && value instanceof String) {
											try {
												value = new Long(Long.parseLong(value.toString()));
												ok = true;
											} catch (Exception ex) {
											}
										}

										if (!ok) {
											conn.rollback();
											return new String[] { "第" + lineNo + "行，'" + nickName + "'的值类型不对，需要" + getDataTypeCaption(type) };
										}
									}

									// BOOL类型的数据插入数据库时需要转义为1，0
									if (value instanceof Boolean) {
										Boolean b = (Boolean) value;
										value = b.booleanValue() ? new Integer(1) : new Integer(0);
									}
								}
							}

							if (StringUtils.isNotEmpty(length)) {
								int len = Integer.parseInt(length);
								if (value != null && value.toString().length() > len) {
									conn.rollback();
									return new String[] { "第" + lineNo + "行，'" + nickName + "'的值超过允许的长度" + length };
								}
							}

							if (StringUtils.isNotEmpty(max)) {
								int imax = Integer.parseInt(max);
								if (value != null && value instanceof Number) {
									if (((Number) value).doubleValue() > imax) {
										conn.rollback();
										return new String[] { "第" + lineNo + "行，'" + nickName + "'的值超过最大值" + max };
									}
								}
							}

							if (StringUtils.isNotEmpty(min)) {
								int imin = Integer.parseInt(min);
								if (value != null && value instanceof Number) {
									if (((Number) value).doubleValue() < imin) {
										conn.rollback();
										return new String[] { "第" + lineNo + "行，'" + nickName + "'的值小于最小值" + min };
									}
								}
							}

							if (StringUtils.isNotEmpty(enumCheck) && value != null) {
								RowSetDynaClass row = db.executeQuery(enumCheck, new Object[] { value });
								row.setNeedClose(false);
								List rows = row.getRows(-1);
								if (rows == null || rows.size() < 1) {
									row.close();
									conn.rollback();
									return new String[] { "第" + lineNo + "行，'" + nickName + "'的枚举值非法" };
								}
								row.close();
							}

							// 找不到值，取缺省值
							if (value == null || "".equals(value)) {
								// 无值，并且是主键，则要insert
								if ("true".equals(primary)) {
									insert = true;

									String repeatLastNo = "";
									if ("true".equals(repeatLast)) {
										repeatLastNo = (String) masterNoMap.get(columnName + "_REPEAT_LAST");
									}

									if (StringUtils.isEmpty(repeatLastNo)) {
										RowSetDynaClass row = db.executeQuery("select " + def + " from dual ");
										row.setNeedClose(false);
										Iterator rowIt = row.getIterator();
										if (rowIt.hasNext()) {
											Object o = rowIt.next();
											value = Database.getProperty(o, row.getDynaProperties()[0].getName());
											keyParams.add(value);
											if ("true".equals(repeatLast)) {
												masterNoMap.put(columnName + "_REPEAT_LAST", value);
												globalRepeatLastNo = value.toString();
											}
										}
										row.close();
									} else {
										value = repeatLastNo;
										keyParams.add(value);
									}
								} else {
									if (def != null) {
										isql = isql.replace("$P{" + columnName + "}", def);
									}
									if ("true".equals(updateFlag) && def != null) {
										usql = usql.replace("$P{" + columnName + "}", def);
									} else {
										// update语句，如果没有找到数据，应该保留原数据
										usql = usql.replace("$P{" + columnName + "}", columnName);
									}
								}
							} else {
								// 如果数据是以#开头，则需要获取默认值
								String key = value.toString().trim();
								if (key.startsWith("#")) {
									// #开头也是需要insert的
									if ("true".equals(primary)) {
										insert = true;
									}
									Object no = masterNoMap.get(columnName + key);
									if (null == no || "".equals(no)) {
										// 默认值不为空，则获取默认值作为NO
										if (StringUtils.isNotEmpty(def)) {
											RowSetDynaClass row = db.executeQuery("select " + def + " from dual ");
											row.setNeedClose(false);
											Iterator rowIt = row.getIterator();
											if (rowIt.hasNext()) {
												Object o = rowIt.next();
												no = Database.getProperty(o, row.getDynaProperties()[0].getName());
												masterNoMap.put(columnName + key, no);
											}
											row.close();
										} else {
											// 否则，使用#1$XXX，使用$后面的数据作为默认值
											no = key.substring(key.indexOf("$") + 1);
										}
									}
									value = no;
								}
								if ("true".equals(primary)) {
									keyParams.add(value);
								}
							}
							// 如果是insertOnly=true的列，表示不能update，所以需要替换
							if ("true".equals(insertOnly)) {
								usql = usql.replace("$P{" + columnName + "}", columnName);
							}
							rowData.put(columnName, value);
						}

						// 判断是否有重复的KEY
						String keyString = "";
						for (int i = 0; i < keyParams.size(); i++) {
							Object v = keyParams.get(i);
							if (v != null) {
								keyString += v;
							}
						}

						if (keyUniqueMap[tableIndex] == null) {
							keyUniqueMap[tableIndex] = new HashMap();
						}
						Object keys = keyUniqueMap[tableIndex].get(keyString);
						if (null != keys) {
							continue;
						} else {
							keyUniqueMap[tableIndex].put(keyString, keyString);
						}

						// 执行前置SQL
						for (int aa = 0; preSqls[tableIndex] != null && aa < preSqls[tableIndex].size(); aa++) {
							String sql = (String) preSqls[tableIndex].get(aa);
							if (StringUtils.isNotEmpty(sql)) {
								sql = DesignQuery.createSQLString(db, sql, rowData);
								db.executeUpdate(sql);
								log.debug("执行前置SQL:" + sql);
							}
						}

						// 执行SQL
						if (insert) {
							isql = DesignQuery.createSQLString(db, isql, rowData);
							db.executeUpdate(isql);
							insert_records[tableIndex]++;
							log.debug("导入Excel的SQL:" + isql);
						} else {
							usql = DesignQuery.createSQLString(db, usql, rowData);
							db.executeUpdate(usql);
							update_records[tableIndex]++;
							log.debug("导入Excel的SQL" + usql);
						}

						// 执行后置SQL
						for (int aa = 0; postSqls[tableIndex] != null && aa < postSqls[tableIndex].size(); aa++) {
							String sql = (String) postSqls[tableIndex].get(aa);
							if (StringUtils.isNotEmpty(sql)) {
								sql = DesignQuery.createSQLString(db, sql, rowData);
								db.executeUpdate(sql);
								log.debug("执行后置SQL:" + sql);
							}
						}

						// 执行数据校验
						for (int aa = 0; validSqls[tableIndex] != null && aa < validSqls[tableIndex].size(); aa++) {
							String sql = (String) validSqls[tableIndex].get(aa);
							String[] sqls = sql.split("\\$\\$");
							sql = sqls[0];
							String validText = "数据非法,Valid Sql=" + sql;
							if (sqls.length > 1) {
								validText = sqls[1];
							}
							if (StringUtils.isNotEmpty(sql)) {
								sql = DesignQuery.createSQLString(db, sql, rowData);
								RowSetDynaClass row = db.executeQuery(sql);
								row.setNeedClose(false);
								List rows = row.getRows(-1);
								if (rows.size() > 0) {
									row.close();
									conn.rollback();
									return new String[] { "第" + lineNo + "行，" + validText };
								}
								row.close();
							}
						}
					}
				}
			}
			conn.commit();
			for (int tableIndex = 0; tableIndex < tableList.size(); tableIndex++) {
				Element table = (Element) tableList.get(tableIndex);
				String title = table.getAttributeValue("nick");
				result[tableIndex] = "本次[" + title + "]导入插入" + insert_records[tableIndex] + "条，更新" + update_records[tableIndex]
						+ "条！";
				log.debug("本次[" + title + "]导入插入" + insert_records[tableIndex] + "条，更新" + update_records[tableIndex] + "条！");
			}
			result[result.length - 1] = globalRepeatLastNo;
			return result;
		} catch (Exception e) {
			e.printStackTrace();
			log.error(e);
			try {
				conn.rollback();
			} catch (SQLException e1) {
			}
		}
		return null;
	}

	/**
	 * 根据规则将Excel内容读取到对象中
	 * 
	 * @param file
	 * @param sheetIndex
	 * @param transfer
	 * @return Iterator
	 */
	public static Collection getCollection(InputStream excelIn, int sheetIndex, String xmlConfig) {
		List list = new ArrayList();
		Map masterMap = new HashMap();

		try {
			Iterator it = ReadExcelUtil.getIterator(excelIn, new int[] { sheetIndex })[0];
			Object[] tables = getTable(it, xmlConfig);
			Element table = (Element) tables[0];
			String[] columns = (String[]) tables[1];

			String detailName = table.getAttributeValue("name");

			Class detailClass = Class.forName(table.getAttributeValue("class"));
			if (detailClass == null) {
				throw new CoreException("指定的明细类" + table.getValue() + "没有定义", null);
			}

			int lineNo = 2;

			// 读取列名
			while (it.hasNext()) {

				List datas = (List) it.next();
				// 判断datas是否都是空白行
				String line = "";
				for (int i = 0; i < datas.size(); i++) {
					Object obj = datas.get(i);
					if (obj != null) {
						line += obj.toString();
					}
				}
				if (null == line || "".equals(line)) {
					continue;
				}

				lineNo++;

				// 针对excel读取的配置文件，一个主bo里可能同时引用了多个子bo，要将在读取数据过程中获取的bo缓存起来 zcl
				// 090712
				Map subBoMap = new HashMap();

				Object detail = detailClass.newInstance();
				// 根据title和columns[i]获取列对应的属性名，设置属性
				Class masterClass = null;
				Object masterObj = null;
				String masterPropertyName = "";
				String masterKey = "";

				for (int i = 0; i < datas.size(); i++) {
					if (i > columns.length - 1) {
						break;
					}
					Element column = (Element) XPath.selectSingleNode(table, "columns/column[@nick='" + columns[i] + "']");
					// 列的标题名称没有找到的情况下，不继续处理 zcl 090712
					if (column != null) {
						String propertyName = column.getAttributeValue("name");
						// 如果主属性还没有处理
						// 这里只支持一个点，也就是一层主
						if (propertyName.indexOf(".") > 0) {
							String[] masterPropertys = propertyName.split("\\.");
							masterPropertyName = masterPropertys[0];
							// 同一条记录中，可能带有2个子对应关系，要判断子对象的名称是否发生变化了 zcl 090712
							if (null == masterObj
									|| !masterObj.getClass().getSimpleName().toLowerCase().equals(masterPropertyName.toLowerCase())) {
								if (subBoMap.containsKey(masterPropertyName)) {
									// 获取已生成的子bo对象 zcl 090712
									Field m = detailClass.getDeclaredField(masterPropertyName);
									masterClass = m.getType();
									masterObj = subBoMap.get(masterPropertyName);
								} else {
									// 提取主的属性名称
									Field m = detailClass.getDeclaredField(masterPropertys[0]);
									// 获得主的类型
									masterClass = m.getType();
									// 创建主
									masterObj = masterClass.newInstance();

									// 缓存子bo对象 zcl 090712
									subBoMap.put(masterPropertyName, masterObj);
								}
							}
							Field f = masterClass.getDeclaredField(masterPropertys[1]);

							f.setAccessible(true);
							// 设置主的属性
							// 因数据格式转换的问题，有可能获取字段内容时出错
							try {
								/*
								 * Object valueObj = null; valueObj = datas.get(i); if (f.getType().getCanonicalName().equals(
								 * "double")) { if (datas.get(i) == null) { valueObj = 0; } else if ("".equals(datas.get(i))) { valueObj
								 * = 0; } else { valueObj = datas.get(i); } f.set(masterObj, valueObj); } else { if (datas.get(i) !=
								 * null)
								 * 
								 * f.set(masterObj, datas.get(i) .toString()); }
								 */
								// 先按照要求进行数据类型转换，确实不能转换的才给错误提示
								getDataValue(f, masterObj, datas.get(i));
							} catch (Exception ex) {
								ex.printStackTrace();
								// 数据类型转换出错的话，要停止导入，并给出相应的提示 zcl 090910
								String msg = getDataTypeCaption(f.getType().getSimpleName());
								throw new Exception("第" + lineNo + "行，“" + column.getAttributeValue("nick") + "” 列 ，必须填入  " + msg + "。");

							}
							masterKey += datas.get(i);
						} else {
							// 设置明细属性
							Field f = detailClass.getDeclaredField(propertyName);
							f.setAccessible(true);
							// 因数据格式转换的问题，有可能获取字段内容时出错，例如数字字段里没有填写内容
							try {
								getDataValue(f, detail, datas.get(i));
							} catch (Exception ex) {
								// 数据类型转换出错的话，要停止导入，并给出相应的提示 zcl 090910
								String msg = getDataTypeCaption(f.getType().getSimpleName());
								throw new Exception("第" + lineNo + "行,“" + column.getAttributeValue("nick") + "” 列 ，必须填入 " + msg + "。");
							}

						}
					}
				}
				// 设置明细的父亲属性
				if (null != masterObj) {
					// 判断父亲是否已经存在
					Object master = masterMap.get(masterKey);
					if (null != master) {
						masterObj = master;
					} else {
						masterMap.put(masterKey, masterObj);
					}

					if (subBoMap.size() == 0) {
						// 将父亲设置到明细
						Field f = detailClass.getDeclaredField(masterPropertyName);
						f.setAccessible(true);
						f.set(detail, masterObj);
					} else {
						// 处理在主BO里引用多个子Bo的情况 zcl 090712
						Set keySet = subBoMap.keySet();
						String[] keys = new String[keySet.size()];
						keySet.toArray(keys);
						for (int i = 0; i < keys.length; i++) {
							Field f = detailClass.getDeclaredField(keys[i]);
							f.setAccessible(true);
							f.set(detail, subBoMap.get(keys[i]));

						}
					}
					// 添加父亲的孩子
					// 并不是所有子bo里都定义了父bo,要控制异常 zcl 090712
					try {
						Field m = masterObj.getClass().getDeclaredField(detailName);
						m.setAccessible(true);
						List details = (List) m.get(masterObj);
						details.add(detail);
						m.set(masterObj, details);
					} catch (Exception ex) {

					}
				}
				list.add(detail);
			}
		} catch (Exception ex) {
			// todo 如果有单引号，要替换，否则会导致页面显示错误
			throw new CoreException(ex.getMessage(), ex);
		}
		// 不太明白这是什么意思 ，先去掉 zcl 090716
		// 读出来的BO都放入到list了，为什么要返回这个？
		// if (masterMap.size() > 0 ) {
		// return masterMap.values();
		// }
		return list;
	}

	// 获取excel的单元格的内容，按照需要的数据格式进行转换，不能转换的才抛出异常
	private static void getDataValue(Field f, Object descObj, Object inputValue) throws Exception {
		String dataTypeName = f.getType().getSimpleName();
		if (dataTypeName == null || "".equals(dataTypeName) || inputValue == null) {
			throw new Exception("未判断的数据类型，请联系系统管理员");
		}
		// 录入内容为空，则抛异常

		// 输入的内容为空时，直接将内容转为
		String value = inputValue.toString();

		try {
			if ("double".equals(dataTypeName)) {
				double result = (new Double(value)).doubleValue();
				f.set(descObj, result);
				return;
			}
			if ("float".equals(dataTypeName)) {
				float result = (new Float(value)).floatValue();
				f.set(descObj, result);
				return;
			}
			if ("int".equals(dataTypeName)) {
				int result = (new Integer(value)).intValue();
				f.set(descObj, result);
				return;
			}
			if ("Integer".equals(dataTypeName)) {
				Integer result = new Integer(value);
				f.set(descObj, result);
				return;
			}
			if ("Double".equals(dataTypeName)) {
				Double result = new Double(value);
				f.set(descObj, result);
				return;
			}
			if ("Float".equals(dataTypeName)) {
				Float result = new Float(value);
				f.set(descObj, result);
				return;
			}
			if ("Date".equals(dataTypeName)) {
				// 看是否能转成日期
				Date result = null;
				if ("Date".equals(inputValue.getClass().getSimpleName())) {
					result = (Date) inputValue;
				} else {
					// 按照2种格式进行转换
					try {
						SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
						result = sf.parse(inputValue.toString());
					} catch (Exception e) {
						try {
							SimpleDateFormat sf = new SimpleDateFormat("yyyyMMdd");
							result = sf.parse(inputValue.toString());
						} catch (Exception e1) {
							throw new Exception("数据类型转换错误");
						}

					}

				}

				f.set(descObj, result);
				return;
			}
			f.set(descObj, inputValue);

		} catch (Exception e) {
			throw new Exception("数据类型转换错误");
		}

	}

	// 获取数据类型的中文名称
	private static String getDataTypeCaption(String dataTypeName) {
		String result = "(未判断的数据类型，请联系系统管理员)";
		if ("double".equals(dataTypeName) || "float".equals(dataTypeName) || "long".equals(dataTypeName)
				|| "number".equals(dataTypeName) || "int".equals(dataTypeName) || "Integer".equals(dataTypeName)
				|| "Double".equals(dataTypeName) || "Float".equals(dataTypeName)) {
			result = "数字";
		}
		if ("Date".equals(dataTypeName) || "date".equals(dataTypeName)) {
			result = "日期";
		}
		if ("String".equals(dataTypeName) || "string".equals(dataTypeName)) {
			result = "文本";
		}
		if ("Boolean".equals(dataTypeName) || "bool".equals(dataTypeName) || "boolean".equals(dataTypeName)) {
			result = "布尔值(填写TRUE或者FALSE。(TRUE代表‘是’,FALSE代表‘否’))";
		}

		return result;
	}

	// 获取配置文件的Element
	private static Element getRootElement(String xmlConfig) {
		InputStream configIn = null;
		try {
			configIn = ReadExcelUtil.class.getResourceAsStream(xmlConfig);
			SAXBuilder sb = new SAXBuilder();
			Document doc = sb.build(configIn);
			Element root = doc.getRootElement();
			return root;
		} catch (Exception ex) {
			throw new CoreException("", ex);
		} finally {
			if (null != configIn) {
				try {
					configIn.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	// 获取Excel中标题对应的XML配置
	private static Object[] getTable(Iterator it, String xmlConfig) {
		try {
			// 读取标题(Excel第一行)
			String title = "";
			if (it.hasNext()) {
				List rows = (List) it.next();
				if (rows != null && rows.size() > 0) {
					// 标题不一定在第1格
					// title = (String) rows.get(0);
					String[] titles = new String[rows.size()];
					rows.toArray(titles);
					for (int j = 0; j < titles.length; j++) {
						if (titles[j] != null && !"".equals(titles[j])) {
							title = titles[j];
							break;
						}
					}
				}
			}

			// 读取列标题(Excel第二行)
			String[] columns = null;
			if (it.hasNext()) {
				List rows = (List) it.next();
				String columnTitle = "";
				for (int i = 0; rows != null && i < rows.size(); i++) {
					columnTitle += ":,:" + (String) rows.get(i);
				}
				if (columnTitle.startsWith(":,:")) {
					columnTitle = columnTitle.substring(3);
				}
				columns = columnTitle.split(":,:");
			}

			Element root = getRootElement(xmlConfig);
			if (null == root) {
				throw new CoreException("Excel的XML配置文件" + xmlConfig + "不存在", null);
			}

			Element table = (Element) XPath.selectSingleNode(root, "/config/tables/table[@nick='" + title + "']");
			if (null == table) {
				throw new CoreException("找不到Excel中标题" + title + "在XML配置中不存在", null);
			}
			return new Object[] { table, columns, root };
		} catch (Exception ex) {
			throw new CoreException("", ex);
		}
	}

	/**
	 * 获取Excel的迭代，用于循环读取Excel内容
	 * 
	 * @param fileName Excel文件名
	 * @param sheetIndex Excel的工作薄索引，从0开始
	 * @return Iterator
	 */
	public static Iterator[] getIterator(InputStream in, int[] sheetIndex) {
		try {
			Iterator[] its = new Iterator[sheetIndex.length];
			HSSFWorkbook wb = WorkbookUtils.openWorkbook(in);
			for (int i = 0; i < sheetIndex.length; i++) {
				HSSFSheet sheet = wb.getSheetAt(sheetIndex[i]);
				its[i] = new ExcelIterator(sheet);
			}
			return its;
		} catch (Exception e) {
			throw new CoreException("", e);
		}
	}

	public static void main(String[] args) {
		/*
		 * InputStream in = null; try { DriverManager.registerDriver(new OracleDriver()); //Connection conn =
		 * DriverManager.getConnection("jdbc:oracle:thin:@192.168.0.180:1521:dstest" , "ds", "ds1234"); Connection conn =
		 * DriverManager.getConnection("jdbc:oracle:thin:@192.168.0.8:1521:ora9i" , "ds", "ds"); //
		 * DriverManager.registerDriver(new com.mysql.jdbc.Driver()); // Connection conn = DriverManager.getConnection( //
		 * "jdbc:mysql://localhost:3306/ky_demo?useUnicode=true&characterEncoding=UTF-8" , // "root", "gzkysm");
		 * 
		 * in = ReadExcelUtil.class.getResourceAsStream("/excel/Book1.xls"); String[] result = saveOrUpdateTable(conn, in,
		 * 4, "/excel/ImportTable.xml"); conn.close(); for (int i = 0; result != null && i < result.length; i++) {
		 * System.out.println(result[i]); } } catch (Exception e) { e.printStackTrace(); } finally { if (null != in) { try {
		 * in.close(); } catch (IOException e) { e.printStackTrace(); } } }
		 */
		Class cla = VendorQuoteMapping.class;
		Field m;
		try {
			m = cla.getDeclaredField("f1");
			System.out.println(m.getType().getSimpleName());
		} catch (SecurityException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (NoSuchFieldException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// double,String,float,int,Integer,Double,Float,Date,

	}
}
