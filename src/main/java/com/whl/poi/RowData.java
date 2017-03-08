package com.whl.poi;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface RowData {
	 String desc() default "";
	 /**
	 * for type to describe the class across the rows;
	 */
	int rowAcross() default 1;
}
