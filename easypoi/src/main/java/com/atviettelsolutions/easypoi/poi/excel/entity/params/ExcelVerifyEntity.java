package com.atviettelsolutions.easypoi.poi.excel.entity.params;

/**
 * Excel check object
 */
public class ExcelVerifyEntity {

    /**
     * interface verification
     *
     * @return
     */
    private boolean interHandler;

    /**
     * empty
     *
     * @return
     */
    private boolean notNull;

    /**
     * is a mobile phone
     *
     * @return
     */
    private boolean isMobile;
    /**
     * is the telephone number
     *
     * @return
     */
    private boolean isTel;

    /**
     * is email
     *
     * @return
     */
    private boolean isEmail;

    /**
     * minimum length
     *
     * @return
     */
    private int minLength;

    /**
     * the maximum length
     *
     * @return
     */
    private int maxLength;

    /**
     * is expressing
     *
     * @return
     */
    private String regex;
    /**
     * is expressing error message
     *
     * @return
     */
    private String regexTip;

    public int getMaxLength() {
        return maxLength;
    }

    public int getMinLength() {
        return minLength;
    }

    public String getRegex() {
        return regex;
    }

    public String getRegexTip() {
        return regexTip;
    }

    public boolean isEmail() {
        return isEmail;
    }

    public boolean isInterHandler() {
        return interHandler;
    }

    public boolean isMobile() {
        return isMobile;
    }

    public boolean isNotNull() {
        return notNull;
    }

    public boolean isTel() {
        return isTel;
    }

    public void setEmail(boolean isEmail) {
        this.isEmail = isEmail;
    }

    public void setInterHandler(boolean interHandler) {
        this.interHandler = interHandler;
    }

    public void setMaxLength(int maxLength) {
        this.maxLength = maxLength;
    }

    public void setMinLength(int minLength) {
        this.minLength = minLength;
    }

    public void setMobile(boolean isMobile) {
        this.isMobile = isMobile;
    }

    public void setNotNull(boolean notNull) {
        this.notNull = notNull;
    }

    public void setRegex(String regex) {
        this.regex = regex;
    }

    public void setRegexTip(String regexTip) {
        this.regexTip = regexTip;
    }

    public void setTel(boolean isTel) {
        this.isTel = isTel;
    }

}
