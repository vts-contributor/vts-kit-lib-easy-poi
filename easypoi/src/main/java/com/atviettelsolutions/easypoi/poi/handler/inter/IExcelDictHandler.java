package com.atviettelsolutions.easypoi.poi.handler.inter;

/**
 * Dictionary translation processing
 */
public interface IExcelDictHandler {

    /**
     * Translate from values to names
     *
     * @param dict
     * @param obj
     * @param name
     * @param value
     * @return
     */
    public String toName(String dict, Object obj, String name, Object value);

    /**
     * Translate from names to values
     *
     * @param dict
     * @param obj
     * @param name
     * @param value
     * @return
     */
    public String toValue(String dict, Object obj, String name, Object value);

}
