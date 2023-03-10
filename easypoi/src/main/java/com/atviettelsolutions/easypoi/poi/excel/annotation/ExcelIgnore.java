package com.atviettelsolutions.easypoi.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Marked as excel to create entities to ignore, resulting in an infinite loop
 * 
 * @author caprocute
 * @date 2013-9-24
 * @version 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelIgnore {

}
