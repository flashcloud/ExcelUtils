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

import org.apache.struts.action.Action;
import org.apache.struts.action.ActionForm;
import org.apache.struts.action.ActionForward;
import org.apache.struts.action.ActionMapping;

@SuppressWarnings("unchecked")
public class TestReportAction extends Action {

  public ActionForward execute(ActionMapping mapping, ActionForm form,
      javax.servlet.http.HttpServletRequest request,
      javax.servlet.http.HttpServletResponse response)
      throws java.lang.Exception {  
    // 准备数据1
    Model model = new Model();
    model.setUser("aaa");
    model.setName("客数码");
    model.setQty(123.234);
    model.setCount(0);
    model.setField1("test");
    model.setYear("2001");

    List details = new ArrayList();
    for (int i = 1; i < 4; i++) {
      Model model1 = new Model();
      model1.setUser("bbbcadff"+(int)(i/2));
      model1.setName("数码客运"+(int)(i/2));
      model1.setQty(909.234+i);
      model1.setCount(i); 
      model1.setYear("200"+(int)(i/3));
      details.add(model1);
    }
    model.setChildren(details);

    Map maps = new LinkedHashMap();
    maps.put("key0", "1");
    maps.put("key1", "数学");
    maps.put("key2", "英语");
    maps.put("key3", "政治");
    maps.put("key4", "历史");    

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
    map111.put("aaa", "标题");
    
    List aList = new ArrayList();
    aList.add("中国");
    aList.add("美国");
    aList.add(map1);    
    aList.add("俄罗斯");
    
    Map context = new HashMap();
    context.put("tranType", "201");
    context.put("printDate", getCurrentDate("yyyy年MM月dd日"));
    context.put("model", model);
    context.put("maps", maps);
    context.put("list", list);
    context.put("index", new Integer(1));
    context.put("key", "key0");
    context.put("where", "数码客运");
    context.put("dd", "Date");
    context.put("patten", "yyyy-MM-dd");
    context.put("width", "2");
    context.put("width1", new Integer(11));
    context.put("title", map111);
    context.put("array", new String[] {"北京","上海","广州"});
    context.put("array_int", new int[] {22,33,44});
    context.put("alist", aList);
    context.put("service", this);
    context.put("stati", TestReportAction.class);
    
    Map context1 = new HashMap();
    context1.put("tranType", "203");
    context1.put("printDate", getCurrentDate("yyyy-MM-dd"));
    context1.put("model", model);
    context1.put("maps", maps);
    context1.put("list", list);
    context1.put("index", new Integer(1));
    context1.put("key", "key0");
    context1.put("where", "数码客运");
    context1.put("dd", "Date");
    context1.put("patten", "yyyy-MM-dd");
    context1.put("width", "2");
    context1.put("width1", new Integer(11));
    context1.put("title", map111);
    context1.put("array", new String[] {"北京","上海","广州"});
    context1.put("array_int", new int[] {22,33,44});
    context1.put("alist", aList);
    context1.put("service", this);
    context1.put("stati", TestReportAction.class);   
    
    Map context2 = new HashMap();
    context2.put("tranType", "200");
    context2.put("printDate", getCurrentDate("yyyy-MM-dd"));
    context2.put("model", model);
    context2.put("maps", maps);
    context2.put("list", list);
    context2.put("index", new Integer(1));
    context2.put("key", "key0");
    context2.put("where", "数码客运");
    context2.put("dd", "Date");
    context2.put("patten", "yyyy-MM-dd");
    context2.put("width", "2");
    context2.put("width1", new Integer(11));
    context2.put("title", map111);
    context2.put("array", new String[] {"北京","上海","广州"});
    context2.put("array_int", new int[] {22,33,44});
    context2.put("alist", aList);
    context2.put("service", this);
    context2.put("stati", TestReportAction.class);     
    
    List lists = new ArrayList();
    lists.add(context);    
    lists.add(context1);    
    lists.add(context2);    
    
    ExcelUtils.addValue("lists", lists);
    
    String config = "/WEB-INF/xls/test.xls";

    response.reset();
    response.setContentType("application/vnd.ms-excel");
    response.setHeader("Content-Disposition", "attachment; filename=\"" + System.currentTimeMillis() + ".xls\"");
    // 输出Excel
    ExcelUtils.export(getServlet().getServletContext(), 
        config, response.getOutputStream());
    return null;
  }

  public String getCurrentDate(String pattern) {
    SimpleDateFormat format = new SimpleDateFormat(pattern);
    return format.format(new Date());
  }
  
  public String getCurrentDate(String pattern, int aaa) {
    SimpleDateFormat format = new SimpleDateFormat(pattern);
    return format.format(new Date())+aaa;
  }
  public static Model getMyModel() {
    Model m = new Model();
    m.setName("aaabbb");
    return m;
  }
  public static Model getMyModel(String a) {
    Model m = new Model();
    m.setName("aaabbb"+a);
    return m;
  }
}
