
package com.viettel.vtskit.easypoi.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel Import Validation
 * 
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelVerify {
	/**
	 * interface verification
	 * 
	 * @return
	 */
	public boolean interHandler() default false;

	/**
	 * is-email
	 * 
	 * @return
	 */
	public boolean isEmail() default false;

	/**
	 * is a mobile phone
	 * 
	 * @return
	 */
	public boolean isMobile() default false;

	/**
	 * is the telephone number
	 * 
	 * @return
	 */
	public boolean isTel() default false;

	/**
	 * the maximum length
	 * 
	 * @return
	 */
	public int maxLength() default -1;

	/**
	 * minimum length
	 * 
	 * @return
	 */
	public int minLength() default -1;

	/**
	 * empty
	 * 
	 * @return
	 */
	public boolean notNull() default false;

	/**
	 * is expressing
	 * 
	 * @return
	 */
	public String regex() default "";

	/**
	 * is expressing, error message
	 * 
	 * @return
	 */
	public String regexTip() default "Data does not meet the validation";

}
