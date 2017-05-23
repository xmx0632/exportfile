package org.xmx0632.exportfile.excel.util;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.xmx0632.exportfile.excel.model.ExcelDataFormatter;
import org.xmx0632.exportfile.excel.model.ExcelSheet;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

public class ExcelUtilsTest {
	String path = "target/test.xlsx";

	@Test
	public void testWriteToFile() throws Exception {
		System.out.println("写Excel");

		List<ExcelTestUser> list = new ArrayList<ExcelTestUser>();
		ExcelTestUser u = createUser("1", "fdsafdsa", 123.23D, new Date(), false, new BigDecimal(123));
		list.add(u);

		u = createUser("222", "fdsafdsa", 123.23D, new Date(), true, new BigDecimal(234));
		list.add(u);

		ExcelDataFormatter edf = createExcelDataFormatter();

		ExcelUtils.writeToFile(list, edf, path, false);

		// List<ExcelTestUser> xx = new ExcelUtils<ExcelTestUser>(new
		// ExcelTestUser()).readFromFile(edf, new File(path));
		// System.out.println(new GsonBuilder().create().toJson(xx));
	}

	@Test
	public void testWriteToFileNSheets() throws Exception {

		System.out.println("写Excel");

		List<ExcelTestUser> list = new ArrayList<ExcelTestUser>();
		ExcelTestUser u = createUser("1", "我是谁", 123.23D, new Date(), false, new BigDecimal(123));
		list.add(u);

		List<ExcelTestUser2> list2 = new ArrayList<ExcelTestUser2>();
		ExcelTestUser2 u2 = createUser2("1", "fdsafdsa");
		list2.add(u2);

		u2 = createUser2("222", "fdsafdsa");
		list2.add(u2);

		ExcelDataFormatter edf = createExcelDataFormatter();

		Map<Class, List<?>> map = Maps.newHashMap();
		map.put(ExcelTestUser.class, list);
		map.put(ExcelTestUser2.class, list2);

		ExcelUtils.writeToFileNSheet(map, edf, path);
	}

	@Test
	public void testWriteToFileNSheet() throws Exception {

		System.out.println("写Excel");

		List<ExcelTestUser> list = new ArrayList<ExcelTestUser>();
		ExcelTestUser u = createUser("1", "我是谁", 123.23D, new Date(), false, new BigDecimal(123));
		list.add(u);

		List<ExcelTestUser2> list2 = new ArrayList<ExcelTestUser2>();
		ExcelTestUser2 u2 = createUser2("1", "fdsafdsa");
		list2.add(u2);

		u2 = createUser2("222", "fdsafdsa");
		list2.add(u2);

		ExcelDataFormatter edf = createExcelDataFormatter();

		List<ExcelSheet> sheetList = Lists.newArrayList();
		ExcelSheet sheet1 = new ExcelSheet();
		sheet1.setTitle("xxx");
		sheet1.setEdf(edf);
		sheet1.setContent(list);
		sheetList.add(sheet1);

		ExcelSheet sheet2 = new ExcelSheet();
		sheet2.setTitle("yyy");
		sheet2.setEdf(edf);
		sheet2.setContent(list2);
		sheetList.add(sheet2);

		ExcelUtils.writeToFileNSheet(sheetList, edf, path);
	}

	private ExcelDataFormatter createExcelDataFormatter() {
		ExcelDataFormatter edf = new ExcelDataFormatter();
		Map<String, String> map = new HashMap<String, String>();
		map.put("true", "已锁定");
		map.put("false", "未锁定");
		edf.set("locked", map);
		return edf;
	}

	private ExcelTestUser createUser(String age, String name, double xx, Date date, boolean locked,
			BigDecimal bigDecimal) {
		ExcelTestUser u = new ExcelTestUser();
		u.setAge(age);
		u.setName(name);
		u.setXx(xx);
		u.setYy(date);
		u.setLocked(locked);
		u.setDb(bigDecimal);
		return u;
	}

	private ExcelTestUser2 createUser2(String age, String name) {
		ExcelTestUser2 u = new ExcelTestUser2();
		u.setAge(age);
		u.setName(name);
		return u;
	}
}
