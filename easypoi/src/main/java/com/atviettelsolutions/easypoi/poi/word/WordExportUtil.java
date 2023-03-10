
package com.atviettelsolutions.easypoi.poi.word;

import com.atviettelsolutions.easypoi.poi.word.parse.ParseWord07;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.Map;

/**
 * Word uses the template export utility class
 *
 * @author caprocute
 * @version 1.0
 */
public final class WordExportUtil {

    private WordExportUtil() {

    }

    /**
     * Parsing the Word 2007 version
     *
     * @param url
     * @param map
     * @return
     */
    public static XWPFDocument exportWord07(String url, Map<String, Object> map) throws Exception {
        return new ParseWord07().parseWord(url, map);
    }

    /**
     * Parsing the Word 2007 version
     *
     * @param map
     * @return
     */
    public static void exportWord07(XWPFDocument document, Map<String, Object> map) throws Exception {
        new ParseWord07().parseWord(document, map);
    }

}
