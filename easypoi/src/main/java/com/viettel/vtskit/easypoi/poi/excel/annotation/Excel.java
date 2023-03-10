package com.viettel.vtskit.easypoi.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel Export basic annotations
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Excel {

    /**
     * Export time setting, if the field is Date type, you don’t need to set it.
     * If the database is string type, you need to set the database format
     */
    public String databaseFormat() default "yyyyMMddHHmmss";

    /**
     * The exported time format, judge whether the date needs to be formatted based on whether this is empty
     */
    public String exportFormat() default "";

    /**
     * Time format, which is equivalent to setting exportFormat and importFormat at the same time
     */
    public String format() default "";

    /**
     * When exporting, the height of each column in excel is in characters, and one Chinese character = 2 characters
     */
    public double height() default 10;

    /**
     * Export type 1 read_old from file, 2 read byte file from database, 3 file address _new, 4 network address The same import is the same
     */
    public int imageType() default 3;

    /**
     * Imported time format, judge whether the date needs to be formatted based on whether this is empty
     */
    public String importFormat() default "";

    /**
     * Text suffix, such as % 90 becomes 90%
     */
    public String suffix() default "";

    /**
     * Whether newline is supported\n
     */
    public boolean isWrap() default true;

    /**
     * Merge cell dependencies, for example, the second column is merged based on the first column, then {1} is fine
     */
    public int[] mergeRely() default {};

    /**
     * Merge cells with the same content vertically
     */
    public boolean mergeVertical() default false;

    /**
     * When exporting, the fields corresponding to the database are mainly for the user to distinguish each field,
     * and there cannot be column names with the same name as the annotation.
     * The export sort is related to the order of the fields that define the annotation.
     * You can use a_id, b_id to determine whether to use it
     */
    public String name();

    /**
     * Do you need to merge cells vertically (used to contain single cells in the list, multiple rows created by merging the list)
     */
    public boolean needMerge() default false;

    /**
     * To display the number, you can use a id, b id to determine the different sorting
     */
    public String orderNum() default "0";

    /**
     * It's worth replacing. The export is {"male_1", "female_0"} and the import is the other way around, so just write a
     */
    public String[] replace() default {};

    /**
     * Import path, if it is a picture, it can be filled in, the default is uploadclassName IconEntity This class corresponds to uploadIcon
     */
    public String savePath() default "upload";

    /**
     * Export type 1 is text, 2 is picture, 3 is function, 4 is number, the default is text
     */
    public int type() default 1;

    /**
     * When exporting, the width unit of each column in excel is characters,
     * one accented characters character = 2 characters,
     * such as the more appropriate length in the content of the column name,
     * such as name column 6 [name generally three characters]
     * gender column 4 [male and female account for 1, But there are two accented characters in the column title] limit 1-255
     */
    public double width() default 10;

    /**
     * Whether to automatically count the data, if it is statistics,
     * add a line of statistics at the end if it is true, and combine all the data.
     * This processing will swallow exceptions, please pay attention to this
     *
     * @return
     */
    public boolean isStatistics() default false;

    /**
     * Method description: Data dictionary table
     * return type： String
     */
    public String dictTable() default "";

    /**
     * Method description: data code
     */
    public String dicCode() default "";

    /**
     * Method description: Data Text
     */
    public String dicText() default "";

    /**
     * Whether the imported data needs to be converted
     * If it is true, you need to add method in pojo: convertset field name (String text)
     *
     * @return
     */
    public boolean importConvert() default false;

    /**
     * Whether the exported data needs to be converted
     * If it is true, you need to add the method in pojo: convertget field name ()
     *
     * @return
     */
    public boolean exportConvert() default false;

    /**
     * Whether the value replacement supports multiple replacements (the default is true,
     * if the database value already contains commas, you need to configure the value to be false)
     */
    public boolean multiReplace() default true;

    /**
     * parent header
     *
     * @return
     */
    String groupName() default "";

    /**
     * Number formatting, the parameter is Pattern, and the object used is Decimal Format
     *
     * @return
     */
    String numFormat() default "";

    /**
     * Do you need to hide this column
     *
     * @return
     */
    public boolean isColumnHidden() default false;

    /**
     * Fix a certain column to solve the problem that is not easy to parse
     *
     * @return
     */
    public int fixedIndex() default -1;

    /**
     * Is this a hyperlink? If it needs to implement the interface to return the object
     *
     * @return
     */
    public boolean isHyperlink() default false;
}
