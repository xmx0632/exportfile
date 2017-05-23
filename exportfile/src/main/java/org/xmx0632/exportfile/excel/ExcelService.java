package org.xmx0632.exportfile.excel;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.Validate;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmx0632.exportfile.excel.model.ExcelDataFormatter;
import org.xmx0632.exportfile.excel.model.ExcelSheet;
import org.xmx0632.exportfile.excel.util.ExcelUtils;

public class ExcelService {

	private static Logger log = LoggerFactory.getLogger(ExcelService.class);

	<T> void dump(String path, Class<T> clazz, ExcelDataFormatter edf) throws Exception {

		// List<ExcelTestUser> xx = new ExcelUtils<ExcelTestUser>(new
		// ExcelTestUser()).readFromFile(edf, new File(path));
		// log.info(new GsonBuilder().create().toJson(xx));
	}

	public <T> void createFile(String path, List<T> list, ExcelDataFormatter edf, boolean noContentFound) {
		Validate.notEmpty(list, "list can't be null");
		log.debug("create excel path:{} list size:{},noContentFound:{}", path, list.size(), noContentFound);
		try {
			File dir = new File(path);
			FileUtils.forceMkdir(dir.getParentFile());
			ExcelUtils.writeToFile(list, edf, path, noContentFound);
		} catch (IOException e) {
			log.error(e.getMessage(), e);
			throw new ExcelException(e.getMessage(), e);
		}
	}
	
	public <T> void createFileWithNSheet(List<ExcelSheet> sheetList, ExcelDataFormatter edf, String path) {
		Validate.notEmpty(sheetList, "list can't be null");
		log.debug("create excel path:{} list size:{},", path, sheetList.size());
		try {
			File dir = new File(path);
			FileUtils.forceMkdir(dir.getParentFile());
			
			ExcelUtils.writeToFileNSheet(sheetList, edf, path);
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			throw new ExcelException(e.getMessage(), e);
		}
	}

}
