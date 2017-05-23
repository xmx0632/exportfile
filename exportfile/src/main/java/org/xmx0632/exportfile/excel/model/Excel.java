package org.xmx0632.exportfile.excel.model;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel注解，用以生成Excel表格文件
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD, ElementType.TYPE })
public @interface Excel {

	// 列名
	String name() default "";

	// 指定日期格式;仅当数据为Date类型时使用
	String pattern() default "yyyy-MM-dd HH:mm:ss";

	// 宽度
	int width() default 20;

	// 忽略该字段
	boolean skip() default false;

}