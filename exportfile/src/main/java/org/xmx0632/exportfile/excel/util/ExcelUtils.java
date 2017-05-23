package org.xmx0632.exportfile.excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmx0632.exportfile.excel.ExcelException;
import org.xmx0632.exportfile.excel.model.Excel;
import org.xmx0632.exportfile.excel.model.ExcelDataFormatter;
import org.xmx0632.exportfile.excel.model.ExcelSheet;


/**
 * Excel导出
 * 
 */
public class ExcelUtils<E> {

	private static Logger log = LoggerFactory.getLogger(ExcelUtils.class);
	private E e;
	private int etimes = 0;

	public ExcelUtils(E e) {
		this.e = e;
	}

	@SuppressWarnings("unchecked")
	private E get() throws InstantiationException, IllegalAccessException {
		return (E) e.getClass().newInstance();
	}

	/**
	 * 将数据写入到EXCEL文档
	 * 
	 * @param list
	 *            数据集合
	 * @param edf
	 *            数据格式化，比如有些数字代表的状态，像是0:女，1：男，或者0：正常，1：锁定，变成可读的文字
	 *            该字段仅仅针对Boolean,Integer两种类型作处理
	 * @param filePath
	 *            文件路径
	 * @param noContentFound 没有查询结果时,只显示标题栏
	 * @throws Exception
	 */
	public static <T> void writeToFile(List<T> list, ExcelDataFormatter edf, String filePath, boolean noContentFound) {
		log.debug("list size:{},filePath:{}", list.size(), filePath);
		// 创建并获取工作簿对象
		Workbook wb = getWorkBook(list, edf, noContentFound);
		log.debug("wb:{}", wb);
		// 写入到文件
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(filePath);
			wb.write(out);
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			throw new ExcelException(e.getMessage(), e);
		} finally {
			IOUtils.closeQuietly(out);
		}
	}

	public static <T> void writeToFileNSheet(List<ExcelSheet> list, ExcelDataFormatter edf, String filePath)
			throws Exception {
		// 创建并获取工作簿对象
		Workbook wb = getWorkBookForNSheet(list, edf);
		// 写入到文件
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(filePath);
			wb.write(out);
			out.close();
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			IOUtils.closeQuietly(out);
			throw e;
		}
	}

	public static <T> void writeToFileNSheet(Map<Class, List<?>> map, ExcelDataFormatter edf, String filePath)
			throws Exception {
		// 创建并获取工作簿对象
		Workbook wb = getWorkBook(map, edf);
		// 写入到文件
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(filePath);
			wb.write(out);
			out.close();
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			IOUtils.closeQuietly(out);
			throw e;
		}
	}

	public static <T> Workbook getWorkBookForNSheet(List<ExcelSheet> sheetList, ExcelDataFormatter edf)
			throws Exception {
		// 创建工作簿
		Workbook wb = new SXSSFWorkbook();

		CreationHelper createHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();

		int sheetIdx = 0;
		for (ExcelSheet excelSheet : sheetList) {
			List<?> list = excelSheet.getContent();
			Field[] fields = Reflections.getClassFieldsAndSuperClassFields(list.get(0).getClass());
			// 创建一个工作表sheet
			Sheet sheet = wb.createSheet();
			wb.setSheetName(sheetIdx, excelSheet.getTitle());

			Font font = getFont(wb);
			XSSFCellStyle titleStyle = getTitleStyle(wb, font);
			// 设置标题
			writeHeader(fields, titleStyle, sheet);
			writeBody(list, edf, createHelper, fields, sheet, cellStyle);
			sheetIdx++;
		}

		return wb;
	}

	public static <T> Workbook getWorkBook(Map<Class, List<?>> map, ExcelDataFormatter edf) throws Exception {
		// 创建工作簿
		Workbook wb = new SXSSFWorkbook();

		CreationHelper createHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();

		Set<Entry<Class, List<?>>> entrySet = map.entrySet();
		for (Entry<Class, List<?>> entry : entrySet) {
			List<?> list = entry.getValue();
			Field[] fields = Reflections.getClassFieldsAndSuperClassFields(list.get(0).getClass());
			// 创建一个工作表sheet
			Sheet sheet = wb.createSheet();

			Font font = getFont(wb);
			XSSFCellStyle titleStyle = getTitleStyle(wb, font);
			// 设置标题
			writeHeader(fields, titleStyle, sheet);
			writeBody(list, edf, createHelper, fields, sheet, cellStyle);
		}

		return wb;
	}

	/**
	 * 获得Workbook对象
	 * 
	 * @param list
	 *            数据集合
	 * @param noContentFound 没有查询结果时,只显示标题栏
	 * @return Workbook
	 * @throws Exception
	 */
	public static <T> Workbook getWorkBook(List<T> list, ExcelDataFormatter edf, boolean noContentFound) {
		// 创建工作簿
		Workbook wb = new SXSSFWorkbook();

		if (list == null || list.size() == 0) {
			return wb;
		}

		CreationHelper createHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();
		log.debug("getClassFieldsAndSuperClassFields");
		Field[] fields = Reflections.getClassFieldsAndSuperClassFields(list.get(0).getClass());
		log.debug("createSheet");
		// 创建一个工作表sheet
		Sheet sheet = wb.createSheet();
		log.debug("getFont");
		Font font = getFont(wb);
		XSSFCellStyle titleStyle = getTitleStyle(wb, font);
		log.debug("writeHeader");
		// 设置标题
		writeHeader(fields, titleStyle, sheet);
		if (!noContentFound) {
			log.debug("writeBody");
			writeBody(list, edf, createHelper, fields, sheet, cellStyle);
		}

		return wb;
	}

	private static XSSFCellStyle getTitleStyle(Workbook wb, Font font) {
		XSSFCellStyle titleStyle = (XSSFCellStyle) wb.createCellStyle();
		titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// 设置前景色
		titleStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(159, 213, 183)));
		titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
		// 设置字体
		titleStyle.setFont(font);
		return titleStyle;
	}

	private static Font getFont(Workbook wb) {
		Font font = wb.createFont();
		font.setColor(HSSFColor.BROWN.index);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		return font;
	}

	private static <T> void writeBody(List<T> list, ExcelDataFormatter edf, CreationHelper createHelper, Field[] fields,
			Sheet sheet, CellStyle cellStyle) {
		try {
			int rowIndex = 1;
			int contentColumnIndex = 0;
			for (T t : list) {
				Row contentRow = sheet.createRow(rowIndex);
				contentColumnIndex = 0;
				Object o = null;
				for (Field field : fields) {

					field.setAccessible(true);

					// 忽略标记skip的字段
					Excel excel = field.getAnnotation(Excel.class);
					if (excel == null || excel.skip() == true) {
						continue;
					}
					// 数据
					Cell cell = contentRow.createCell(contentColumnIndex);

					o = field.get(t);
					// 如果数据为空，跳过
					if (o == null) {
						contentColumnIndex++;
						continue;
					}

					// 处理日期类型
					if (o instanceof Date) {
						String pattern = excel.pattern();
						log.debug("pattern:{}", pattern);
//						cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(pattern));
//						cell.setCellStyle(cellStyle);
						SimpleDateFormat sdf = new SimpleDateFormat(pattern);
						String dateString = sdf.format(field.get(t));
						cell.setCellValue(dateString);
					} else if (o instanceof Double || o instanceof Float) {
						cell.setCellValue((Double) field.get(t));
					} else if (o instanceof Boolean) {
						Boolean bool = (Boolean) field.get(t);
						if (edf == null) {
							cell.setCellValue(bool);
						} else {
							log.debug("field.getName:{}", field.getName());
							Map<String, String> map = edf.get(field.getName());
							log.debug("map:{}", map);
							if (map == null) {
								cell.setCellValue(bool);
							} else {
								cell.setCellValue(map.get(bool.toString().toLowerCase()));
							}
						}

					} else if (o instanceof Integer) {

						Integer intValue = (Integer) field.get(t);

						if (edf == null) {
							cell.setCellValue(intValue);
						} else {
							Map<String, String> map = edf.get(field.getName());
							if (map == null) {
								cell.setCellValue(intValue);
							} else {
								cell.setCellValue(map.get(intValue.toString()));
							}
						}
					}else if (o instanceof String) {

						String stringValue = (String) field.get(t);

						if (edf == null) {
							cell.setCellValue(stringValue);
						} else {
							Map<String, String> map = edf.get(field.getName());
							if (map == null) {
								cell.setCellValue(stringValue);
							} else {
								cell.setCellValue(map.get(stringValue.toString()));
							}
						}
					}
					else {
						cell.setCellValue(field.get(t).toString());
					}

					contentColumnIndex++;
				}

				rowIndex++;
			}
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			throw new ExcelException(e.getMessage(), e);
		}
	}

	private static void writeHeader(Field[] fields, XSSFCellStyle titleStyle, Sheet sheet) {
		Row headerRow = sheet.createRow(0);
		int columnIndex = 0;
		for (Field field : fields) {
			field.setAccessible(true);
			Excel excel = field.getAnnotation(Excel.class);
			if (excel == null || excel.skip()) {
				continue;
			}
			// 列宽注意乘256
			sheet.setColumnWidth(columnIndex, excel.width() * 256);
			// 写入标题
			Cell cell = headerRow.createCell(columnIndex);
			cell.setCellStyle(titleStyle);
			cell.setCellValue(excel.name());

			columnIndex++;
		}
	}

	/**
	 * 本方法只支持读取单个sheet
	 * 
	 * 从文件读取数据，最好是所有的单元格都是文本格式，日期格式要求yyyy-MM-dd HH:mm:ss,布尔类型0：真，1：假
	 * 
	 * @param edf
	 *            数据格式化
	 * 
	 * @param file
	 *            Excel文件，支持xlsx后缀，xls的没写，基本一样
	 * @return
	 * @throws Exception
	 */
	public List<E> readFromFile(ExcelDataFormatter edf, File file) throws Exception {
		Field[] fields = Reflections.getClassFieldsAndSuperClassFields(e.getClass());

		Map<String, String> textToKey = new HashMap<String, String>();

		Excel _excel = null;
		for (Field field : fields) {
			_excel = field.getAnnotation(Excel.class);
			if (_excel == null || _excel.skip() == true) {
				continue;
			}
			textToKey.put(_excel.name(), field.getName());
		}

		InputStream is = new FileInputStream(file);

		Workbook wb = new XSSFWorkbook(is);

		Sheet sheet = wb.getSheetAt(0);
		Row title = sheet.getRow(0);
		// 标题数组，后面用到，根据索引去标题名称，通过标题名称去字段名称用到 textToKey
		String[] titles = new String[title.getPhysicalNumberOfCells()];
		for (int i = 0; i < title.getPhysicalNumberOfCells(); i++) {
			titles[i] = title.getCell(i).getStringCellValue();
		}

		List<E> list = new ArrayList<E>();

		E e = null;

		int rowIndex = 0;
		int columnCount = titles.length;
		Cell cell = null;
		Row row = null;

		for (Iterator<Row> it = sheet.rowIterator(); it.hasNext();) {

			row = it.next();
			if (rowIndex++ == 0) {
				continue;
			}

			if (row == null) {
				break;
			}

			e = get();

			for (int i = 0; i < columnCount; i++) {
				cell = row.getCell(i);
				etimes = 0;
				readCellContent(textToKey.get(titles[i]), fields, cell, e, edf);
			}
			list.add(e);
		}
		return list;
	}

	/**
	 * 从单元格读取数据，根据不同的数据类型，使用不同的方式读取<br>
	 * 有时候POI自作聪明，经常和我们期待的数据格式不一样，会报异常，<br>
	 * 我们这里采取强硬的方式<br>
	 * 使用各种方法，知道尝试到读到数据为止，然后根据Bean的数据类型，进行相应的转换<br>
	 * 如果尝试完了（总共7次），还是不能得到数据，那么抛个异常出来，没办法了
	 * 
	 * @param key
	 *            当前单元格对应的Bean字段
	 * @param fields
	 *            Bean所有的字段数组
	 * @param cell
	 *            单元格对象
	 * @param e
	 * @throws Exception
	 */
	private void readCellContent(String key, Field[] fields, Cell cell, E e, ExcelDataFormatter edf) throws Exception {

		Object o = null;
		try {
			switch (cell.getCellType()) {
			case XSSFCell.CELL_TYPE_BOOLEAN:
				o = cell.getBooleanCellValue();
				break;
			case XSSFCell.CELL_TYPE_NUMERIC:
				o = cell.getNumericCellValue();
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					o = DateUtil.getJavaDate(cell.getNumericCellValue());
				}
				break;
			case XSSFCell.CELL_TYPE_STRING:
				o = cell.getStringCellValue();
				break;
			case XSSFCell.CELL_TYPE_ERROR:
				o = cell.getErrorCellValue();
				break;
			case XSSFCell.CELL_TYPE_BLANK:
				o = null;
				break;
			case XSSFCell.CELL_TYPE_FORMULA:
				o = cell.getCellFormula();
				break;
			default:
				o = null;
				break;
			}

			if (o == null)
				return;

			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			for (Field field : fields) {
				field.setAccessible(true);
				if (field.getName().equals(key)) {
					Boolean bool = true;
					Map<String, String> map = null;
					if (edf == null) {
						bool = false;
					} else {
						map = edf.get(field.getName());
						if (map == null) {
							bool = false;
						}
					}

					if (field.getType().equals(Date.class)) {
						if (o.getClass().equals(Date.class)) {
							field.set(e, o);
						} else {
							field.set(e, sdf.parse(o.toString()));
						}
					} else if (field.getType().equals(String.class)) {
						if (o.getClass().equals(String.class)) {
							field.set(e, o);
						} else {
							field.set(e, o.toString());
						}
					} else if (field.getType().equals(Long.class)) {
						if (o.getClass().equals(Long.class)) {
							field.set(e, o);
						} else {
							field.set(e, Long.parseLong(o.toString()));
						}
					} else if (field.getType().equals(Integer.class)) {
						if (o.getClass().equals(Integer.class)) {
							field.set(e, o);
						} else {
							// 检查是否需要转换
							if (bool) {
								field.set(e, map.get(o.toString()) != null ? Integer.parseInt(map.get(o.toString()))
										: Integer.parseInt(o.toString()));
							} else {
								field.set(e, Integer.parseInt(o.toString()));
							}

						}
					} else if (field.getType().equals(BigDecimal.class)) {
						if (o.getClass().equals(BigDecimal.class)) {
							field.set(e, o);
						} else {
							field.set(e, BigDecimal.valueOf(Double.parseDouble(o.toString())));
						}
					} else if (field.getType().equals(Boolean.class)) {
						if (o.getClass().equals(Boolean.class)) {
							field.set(e, o);
						} else {
							// 检查是否需要转换
							if (bool) {
								field.set(e, map.get(o.toString()) != null ? Boolean.parseBoolean(map.get(o.toString()))
										: Boolean.parseBoolean(o.toString()));
							} else {
								field.set(e, Boolean.parseBoolean(o.toString()));
							}
						}
					} else if (field.getType().equals(Float.class)) {
						if (o.getClass().equals(Float.class)) {
							field.set(e, o);
						} else {
							field.set(e, Float.parseFloat(o.toString()));
						}
					} else if (field.getType().equals(Double.class)) {
						if (o.getClass().equals(Double.class)) {
							field.set(e, o);
						} else {
							field.set(e, Double.parseDouble(o.toString()));
						}

					}

				}
			}

		} catch (Exception ex) {
			log.error(ex.getMessage(), ex);
			// 如果还是读到的数据格式还是不对，只能放弃了
			if (etimes > 7) {
				throw ex;
			}
			etimes++;
			if (o == null) {
				readCellContent(key, fields, cell, e, edf);
			}
		}
	}

}