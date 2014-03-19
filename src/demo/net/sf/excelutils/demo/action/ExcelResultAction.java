package net.sf.excelutils.demo.action;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import net.sf.excelutils.ExcelUtils;
import net.sf.excelutils.demo.bo.Model;

import com.opensymphony.xwork.ActionSupport;


@SuppressWarnings("unchecked")
public class ExcelResultAction extends ActionSupport {

	private static final long serialVersionUID = 1L;

	private Model model;

	private List list;

	private Date printDate;

	private Map maps;

	private Integer index;

	private String key;

	private Object service;

	private Object stati;

	public String execute() {
		model = new Model();
		model.setUser("aaa");
		model.setName("客数码");
		model.setQty(123.234);
		model.setCount(0);
		model.setField1("test");
		model.setYear("2001");

		List details = new ArrayList();
		for (int i = 1; i < 4; i++) {
			Model model1 = new Model();
			model1.setUser("bbbcadff" + (int) (i / 2));
			model1.setName("数码客运" + (int) (i / 2));
			model1.setQty(909.234 + i);
			model1.setCount(i);
			model1.setYear("200" + (int) (i / 3));
			details.add(model1);
		}
		model.setChildren(details);

		maps = new LinkedHashMap();
		maps.put("key0", "1");
		maps.put("key1", "数学");
		maps.put("key2", "英语");

		list = new ArrayList();
		Map map0 = new LinkedHashMap();
		map0.put("key0", new Integer(90));
		map0.put("key1", new Integer(92));
		map0.put("key2", new Integer(89));
		list.add(map0);
		Map map1 = new LinkedHashMap();
		map1.put("key0", new Integer(95));
		map1.put("key1", new Integer(90));
		map1.put("key2", new Double(80.03));
		list.add(map1);

		index = new Integer(1);
		key = "key0";

		service = this;
		stati = ExcelReportAction.class;

		List aList = new ArrayList();
		Map m1 = new HashMap();
		m1.put("a", "中国11");
		m1.put("b", "美国11");
		aList.add(m1);
		Map m2 = new HashMap();
		m2.put("a", "中国12");
		m2.put("b", "美国11");
		aList.add(m2);
		Map m214 = new HashMap();
		m214.put("a", "中国24");
		m214.put("b", "美国24");
		aList.add(m214);
		Map m211 = new HashMap();
		m211.put("a", "中国21");
		m211.put("b", "美国21");
		aList.add(m211);
		Map m221 = new HashMap();
		m221.put("a", "中国22");
		m221.put("b", "美国21");
		aList.add(m221);

		List aList1 = new ArrayList();
		Map m21 = new HashMap();
		m21.put("a", "中国21");
		m21.put("b", "美国21");
		aList1.add(m21);
		Map m22 = new HashMap();
		m22.put("a", "中国22");
		m22.put("b", "美国22");
		aList1.add(m22);

		List sheets = new ArrayList();
		Map sheet0 = new HashMap();
		sheet0.put("name", "页签0");
		sheet0.put("value", "SHEET测试0");
		sheet0.put("list", aList);
		sheets.add(sheet0);
		Map sheet1 = new HashMap();
		sheet1.put("name", "页签1");
		sheet1.put("value", "SHEET测试1");
		sheet1.put("list", aList1);
		sheets.add(sheet1);

		ExcelUtils.addValue("where", "数码客运");
		ExcelUtils.addValue("dd", "Date");
		ExcelUtils.addValue("patten", "yyyy-MM-dd");
		ExcelUtils.addValue("width", "2");
		ExcelUtils.addValue("width1", new Integer(11));
		ExcelUtils.addValue("title", new String[] { "标题" });
		ExcelUtils.addValue("array", new String[] { "北京", "上海", "广州" });
		ExcelUtils.addValue("array_int", new int[] { 22, 33, 44 });
		ExcelUtils.addValue("alist", aList);
		ExcelUtils.addValue("sheets", sheets);

		printDate = new Date();

		return SUCCESS;
	}

	public String getCurrentDate(String pattern) {
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		return format.format(new Date());
	}

	public String getCurrentDate(String pattern, int aaa) {
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		return format.format(new Date()) + aaa;
	}

	/**
	 * @return Returns the index.
	 */
	public Integer getIndex() {
		return index;
	}

	/**
	 * @param index The index to set.
	 */
	public void setIndex(Integer index) {
		this.index = index;
	}

	/**
	 * @return Returns the key.
	 */
	public String getKey() {
		return key;
	}

	/**
	 * @param key The key to set.
	 */
	public void setKey(String key) {
		this.key = key;
	}

	/**
	 * @return Returns the list.
	 */
	public List getList() {
		return list;
	}

	/**
	 * @param list The list to set.
	 */
	public void setList(List list) {
		this.list = list;
	}

	/**
	 * @return Returns the maps.
	 */
	public Map getMaps() {
		return maps;
	}

	/**
	 * @param maps The maps to set.
	 */
	public void setMaps(Map maps) {
		this.maps = maps;
	}

	/**
	 * @return Returns the model.
	 */
	public Model getModel() {
		return model;
	}

	/**
	 * @param model The model to set.
	 */
	public void setModel(Model model) {
		this.model = model;
	}

	/**
	 * @return Returns the printDate.
	 */
	public Date getPrintDate() {
		return printDate;
	}

	/**
	 * @param printDate The printDate to set.
	 */
	public void setPrintDate(Date printDate) {
		this.printDate = printDate;
	}

	public Object getService() {
		return service;
	}

	public void setService(Object service) {
		this.service = service;
	}

	public Object getStati() {
		return stati;
	}

	public void setStati(Object action) {
		this.stati = action;
	}
}
