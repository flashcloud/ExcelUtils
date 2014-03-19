package net.sf.excelutils.demo.action;

import java.io.ByteArrayOutputStream;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import net.sf.excelutils.ExcelUtils;
import net.sf.excelutils.demo.bo.Model;

import org.apache.struts.action.Action;
import org.apache.struts.action.ActionForm;
import org.apache.struts.action.ActionForward;
import org.apache.struts.action.ActionMapping;

@SuppressWarnings("unchecked")
public class ExcelReportAction extends Action {

	// 从图片里面得到字节数组
	private static byte[] getImageData() {
		try {
			URL url = ExcelReportAction.class.getResource("/image.jpg");
			ByteArrayOutputStream bout = new ByteArrayOutputStream();
			ImageIO.write(ImageIO.read(url), "JPG", bout);
			return bout.toByteArray();
		} catch (Exception exe) {
			exe.printStackTrace();
			return null;
		}
	}

	public ActionForward execute(ActionMapping mapping, ActionForm form, javax.servlet.http.HttpServletRequest request,
			javax.servlet.http.HttpServletResponse response) throws java.lang.Exception {
		// 准备数据
		Model model = new Model();
		model.setUser("aaa");
		model.setName("测试用户");
		model.setQty(123.234);
		model.setCount(0);
		model.setField1("test");
		model.setYear("2001");

		List details = new ArrayList();
		for (int i = 1; i < 4; i++) {
			Model model1 = new Model();
			model1.setUser("bbbcadff" + (int) (i / 2));
			model1.setName("测试" + (int) (i / 2));
			model1.setQty(909.234 + i);
			model1.setCount(i);
			model1.setYear("200" + (int) (i / 3));
			details.add(model1);
		}
		model.setChildren(details);

		Map maps = new LinkedHashMap();
		maps.put("key0", "1");
		maps.put("key1", "数学");
		maps.put("key2", "语文");
		maps.put("key3", "政治");
		maps.put("key4", "历史");

		List keys = new ArrayList();
		keys.add("key4");
		keys.add("key2");

		List list = new ArrayList();
		Map map0 = new LinkedHashMap();
		map0.put("key0", new Integer(90));
		map0.put("key1", new Integer(92));
		map0.put("key2", new Integer(89));
		map0.put("key3", new Integer(69));
		map0.put("key4", new Integer(72));
		list.add(map0);
		Map map1 = new LinkedHashMap();
		map1.put("key0", new Integer(95));
		map1.put("key1", new Integer(90));
		map1.put("key2", new Double(80.03));
		map1.put("key3", new Integer(64));
		map1.put("key4", new Integer(77));
		list.add(map1);

		Map map111 = new LinkedHashMap();
		map111.put("aaa", "abcd");

		List aList = new ArrayList();
		aList.add("月");
		aList.add("9");
		aList.add(map1);
		aList.add("aa");

		List sheets = new ArrayList();
		Map sheet0 = new HashMap();
		sheet0.put("name", "页签0");
		sheet0.put("value", "SHEET测试0");
		sheet0.put("list", list);
		sheets.add(sheet0);
		Map sheet1 = new HashMap();
		sheet1.put("name", "页签1");
		sheet1.put("value", "SHEET测试1");
		sheet1.put("list", list);
		sheets.add(sheet1);

		ExcelUtils.addValue("printDate", getCurrentDate("yyyyMMdd"));
		ExcelUtils.addValue("field", "name");
		ExcelUtils.addValue("model", model);
		ExcelUtils.addValue("maps", maps);
		ExcelUtils.addValue("keys", keys);
		ExcelUtils.addValue("list", list);
		ExcelUtils.addValue("index", new Integer(1));
		ExcelUtils.addValue("key", "key0");
		ExcelUtils.addValue("where", "数码客运");
		ExcelUtils.addValue("dd", "Date");
		ExcelUtils.addValue("patten", "yyyy-MM-dd");
		ExcelUtils.addValue("width", "2");
		ExcelUtils.addValue("width1", new Integer(11));
		ExcelUtils.addValue("title", map111);
		ExcelUtils.addValue("array", new String[] { "北京", "上海", "广州" });
		ExcelUtils.addValue("array_int", new int[] { 22, 33, 44 });
		ExcelUtils.addValue("alist", aList);
		ExcelUtils.addService("service", this);
		ExcelUtils.addService("stati", ExcelReportAction.class);
		ExcelUtils.addValue("sheets", sheets);
		ExcelUtils.addValue("imageData", getImageData());

		String config = "/WEB-INF/xls/demo.xls";

		response.reset();
		response.setContentType("application/vnd.ms-excel");
		// Excel
		ExcelUtils.export(getServlet().getServletContext(), config, response.getOutputStream());
		return null;
	}

	public String getCurrentDate(String pattern) {
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		return format.format(new Date());
	}

	public String getCurrentDate(String pattern, int aaa) {
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		return format.format(new Date()) + aaa;
	}

	public static Model getMyModel() {
		Model m = new Model();
		m.setName("aaabbb");
		return m;
	}

	public static Model getMyModel(String a) {
		Model m = new Model();
		m.setName("aaabbb" + a);
		return m;
	}
}
