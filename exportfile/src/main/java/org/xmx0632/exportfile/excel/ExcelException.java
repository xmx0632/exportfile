package org.xmx0632.exportfile.excel;

/**
 * 业务层异常
 *
 */
public class ExcelException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	public ExcelException() {
		super();
	}

	public ExcelException(String message) {
		super(message);
	}

	public ExcelException(String message, Exception e) {
		super(message,e);
	}

}
