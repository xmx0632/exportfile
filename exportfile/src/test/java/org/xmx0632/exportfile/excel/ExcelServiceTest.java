package org.xmx0632.exportfile.excel;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;
import org.xmx0632.exportfile.excel.model.ExcelDataFormatter;
import org.xmx0632.exportfile.excel.model.ExcelSheet;
import org.xmx0632.exportfile.excel.util.ExcelTestUser;
import org.xmx0632.exportfile.excel.util.ExcelTestUser2;

import com.google.common.collect.Lists;

public class ExcelServiceTest {

	private ExcelService excelService = new ExcelService();

	@Test
	public void testDump() throws InvalidFormatException, IOException {
	}

	@Test
	public void testCreate() throws Exception {
		String path = "target/test.xlsx";
		List<ExcelTestUser> list = new ArrayList<ExcelTestUser>();
		ExcelTestUser u = createUser("1", "fdsafdsa", 123.23D, new Date(), false, new BigDecimal(123));
		list.add(u);

		u = createUser("222", "fdsafdsa", 123.23D, new Date(), true, new BigDecimal(234));
		list.add(u);

		u = createUser("123", "fdsafdsa", 123.23D, new Date(), false, new BigDecimal(2344));
		list.add(u);

		ExcelDataFormatter edf = createExcelDataFormatter();

		excelService.createFile(path, list, edf, false);
	}

	@Test
	public void testCreateNoBody() throws Exception {
		String path = "target/test.xlsx";
		List<ExcelTestUser> list = new ArrayList<ExcelTestUser>();
		ExcelTestUser u = createUser("1", "fdsafdsa", 123.23D, new Date(), false, new BigDecimal(123));
		list.add(u);

		u = createUser("222", "fdsafdsa", 123.23D, new Date(), true, new BigDecimal(234));
		list.add(u);

		u = createUser("123", "fdsafdsa", 123.23D, new Date(), false, new BigDecimal(2344));
		list.add(u);

		ExcelDataFormatter edf = createExcelDataFormatter();

		excelService.createFile(path, list, edf, true);
	}

	private ExcelDataFormatter createExcelDataFormatter() {
		ExcelDataFormatter edf = new ExcelDataFormatter();
		Map<String, String> map = new HashMap<String, String>();
		map.put("真", "true");
		map.put("假", "false");
		edf.set("locked", map);
		return edf;
	}

	private ExcelTestUser createUser(String age, String name, double xx, Date date, boolean locked,
			BigDecimal bigDecimal) {
		ExcelTestUser u;
		u = new ExcelTestUser();
		u.setAge(age);
		u.setName(name);
		u.setXx(xx);
		u.setYy(date);
		u.setLocked(locked);
		u.setDb(bigDecimal);
		return u;
	}

	@Test
	public void testCreateNSheet() throws Exception {

		System.out.println("写Excel");
		String path = "target/testNsheet1.xlsx";

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
		sheet1.setTitle("第一个表格");
		sheet1.setEdf(edf);
		sheet1.setContent(list);
		sheetList.add(sheet1);

		ExcelSheet sheet2 = new ExcelSheet();
		sheet2.setTitle("第二个表格");
		sheet2.setEdf(edf);
		sheet2.setContent(list2);
		sheetList.add(sheet2);

		excelService.createFileWithNSheet(sheetList, edf, path);

	}

	private ExcelTestUser2 createUser2(String age, String name) {
		ExcelTestUser2 u = new ExcelTestUser2();
		u.setAge(age);
		u.setName(name);
		return u;
	}
}
