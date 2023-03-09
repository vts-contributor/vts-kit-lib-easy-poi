package com.viettel.vtskit.easypoi.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Is the tag exported to excel as an entity class?
 * 
 * @author caprocute
 * 
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelEntity {

	/**
	 * Define the excel export ID to limit the export fields, and deal with the situation where a class corresponds to multiple different names
	 */
	public String id() default "";

	/**
	 * When exporting, the fields corresponding to the database are mainly for the user to distinguish each field,
	 * and the column name when exporting cannot have the same name as the annocation
	 * The export sort is related to the order of the fields that define the annotation. You can use a_id, b_id to determine whether to use
	 */
	public String name() default "";

	/**
	 * When exporting, whether to display the text corresponding to the name
	 * @return
	 */
	boolean show() default false;
}
