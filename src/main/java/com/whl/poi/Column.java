package com.whl.poi;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.METHOD,ElementType.FIELD,ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Column {
	 int colIndex() ;
	 int rowIndex() default 0;
	 String desc() default "";
	 String format() default "";
	 boolean isAcross() default false;
	 /**
	 * for type to describe the class across the rows;
	 */
	int rowAcross() default 1;
	dealType dealType() default dealType.CONTINUE_CELL;
	 
	 
	 public enum dealType {
	        /**
	         *  <p>If field and cell do not match, set as null and continue processing other cell</p>
	         */
	        CONTINUE_CELL,
	        
	        /**
	         *  <p>If field and cell do not match, set as null and continue processing other row</p>
	         */
	        CONTINUE_ROW,
	        
	        /**
	         *  <p>If field and cell do not match, abort process the sheet</p>
	         */
	        ABORT,

	    }
}
