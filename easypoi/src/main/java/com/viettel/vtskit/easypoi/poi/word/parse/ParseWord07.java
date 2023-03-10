
package com.viettel.vtskit.easypoi.poi.word.parse;

import com.viettel.vtskit.easypoi.poi.common.PoiPublicUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import com.viettel.vtskit.easypoi.poi.cache.WordCache;
import com.viettel.vtskit.easypoi.poi.word.entity.MyXWPFDocument;
import com.viettel.vtskit.easypoi.poi.word.entity.WordImageEntity;
import com.viettel.vtskit.easypoi.poi.word.entity.params.ExcelListEntity;
import com.viettel.vtskit.easypoi.poi.word.parse.excel.ExcelEntityParse;
import com.viettel.vtskit.easypoi.poi.word.parse.excel.ExcelMapParse;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Parse version 07 of Word, replace text, generate tables, generate pictures
 *
 * @author caprocute
 * @version 1.0
 * @date 2013-11-16
 */
@SuppressWarnings({"unchecked", "rawtypes"})
public class ParseWord07 {

    private static final Logger LOGGER = LoggerFactory.getLogger(ParseWord07.class);

    /**
     * Add a picture
     *
     * @param obj
     * @param currentRun
     * @throws Exception
     * @author caprocute
     */
    private void addAnImage(WordImageEntity obj, XWPFRun currentRun) throws Exception {
        Object[] isAndType = PoiPublicUtil.getIsAndType(obj);
        String picId;
        try {
            picId = currentRun.getParagraph().getDocument().addPictureData((byte[]) isAndType[0], (Integer) isAndType[1]);
            ((MyXWPFDocument) currentRun.getParagraph().getDocument()).createPicture(currentRun, picId,
                    currentRun.getParagraph().getDocument().getNextPicNameNumber((Integer) isAndType[1]), obj.getWidth(), obj.getHeight());

        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
        }

    }

    /**
     * Change the value based on the condition
     *
     * @param map
     * @author caprocute
     */
    private void changeValues(XWPFParagraph paragraph, XWPFRun currentRun, String currentText, List<Integer> runIndex, Map<String, Object> map) throws Exception {
        Object obj = PoiPublicUtil.getRealValue(currentText, map);
        if (obj instanceof WordImageEntity) {// If it is an image, it is set as an image
            currentRun.setText("", 0);
            addAnImage((WordImageEntity) obj, currentRun);
        } else {
            currentText = obj.toString();
            currentRun.setText(currentText, 0);
        }
        for (int k = 0; k < runIndex.size(); k++) {
            paragraph.getRuns().get(runIndex.get(k)).setText("", 0);
        }
        runIndex.clear();
    }

    /**
     * Determine whether it is an iterative output
     *
     * @return
     * @throws Exception
     * @author caprocute
     * @date 2013-11-18
     */
    private Object checkThisTableIsNeedIterator(XWPFTableCell cell, Map<String, Object> map) throws Exception {
        String text = cell.getText().trim();
        if (text != null && text.startsWith("{{") && text.indexOf("$fe:") != -1) {
            return PoiPublicUtil.getRealValue(text.replace("$fe:", "").trim(), map);
        }
        return null;
    }

    /**
     * Parse all text
     *
     * @param paragraphs
     * @param map
     * @author caprocute
     */
    private void parseAllParagraphic(List<XWPFParagraph> paragraphs, Map<String, Object> map) throws Exception {
        XWPFParagraph paragraph;
        for (int i = 0; i < paragraphs.size(); i++) {
            paragraph = paragraphs.get(i);
            if (paragraph.getText().indexOf("{{") != -1) {
                parseThisParagraph(paragraph, map);
            }

        }

    }

    /**
     * Parse this paragraph
     *
     * @param paragraph
     * @param map
     * @author caprocute
     */
    private void parseThisParagraph(XWPFParagraph paragraph, Map<String, Object> map) throws Exception {
        XWPFRun run;
        XWPFRun currentRun = null;// The first run you get, used to set the value, can save the format
        String currentText = "";// Holds the current text
        String text;
        Boolean isfinde = false;// Determine if you have encountered {{
        List<Integer> runIndex = new ArrayList<Integer>();// Store the runs you encounter and empty them
        for (int i = 0; i < paragraph.getRuns().size(); i++) {
            run = paragraph.getRuns().get(i);
            text = run.getText(0);
            if (StringUtils.isEmpty(text)) {
                continue;
            }// If empty, or "" this continues to skip the loop
            if (isfinde) {
                currentText += text;
                if (currentText.indexOf("{{") == -1) {
                    isfinde = false;
                    runIndex.clear();
                } else {
                    runIndex.add(i);
                }
                if (currentText.indexOf("}}") != -1) {
                    changeValues(paragraph, currentRun, currentText, runIndex, map);
                    currentText = "";
                    isfinde = false;
                }
            } else if (text.indexOf("{") >= 0) {//
                currentText = text;
                isfinde = true;
                currentRun = run;
            } else {
                currentText = "";
            }
            if (currentText.indexOf("}}") != -1) {
                changeValues(paragraph, currentRun, currentText, runIndex, map);
                isfinde = false;
            }
        }

    }

    private void parseThisRow(List<XWPFTableCell> cells, Map<String, Object> map) throws Exception {
        for (XWPFTableCell cell : cells) {
            parseAllParagraphic(cell.getParagraphs(), map);
        }
    }

    /**
     * Parse this table
     *
     * @param table
     * @param map
     * @author caprocute
     */
    private void parseThisTable(XWPFTable table, Map<String, Object> map) throws Exception {
        XWPFTableRow row;
        List<XWPFTableCell> cells;
        Object listobj;
        for (int i = 0; i < table.getNumberOfRows(); i++) {
            row = table.getRow(i);
            cells = row.getTableCells();
            listobj = checkThisTableIsNeedIterator(cells.get(0), map);
            if (listobj == null) {
                parseThisRow(cells, map);
            } else if (listobj instanceof ExcelListEntity) {
                new ExcelEntityParse().parseNextRowAndAddRow(table, i, (ExcelListEntity) listobj);
                i = i + ((ExcelListEntity) listobj).getList().size() - 1;
            } else {
                ExcelMapParse.parseNextRowAndAddRow(table, i, (List) listobj);
                i = i + ((List) listobj).size() - 1;
            }
        }
    }

    /**
     * Parse version 07 of Word and assign values
     *
     * @return
     * @throws Exception
     * @author caprocute
     */
    public XWPFDocument parseWord(String url, Map<String, Object> map) throws Exception {
        MyXWPFDocument doc = WordCache.getXWPFDocumen(url);
        parseWordSetValue(doc, map);
        return doc;
    }

    /**
     * Parse version 07 of Word and assign values
     *
     * @return
     * @throws Exception
     * @author caprocute
     * @date 2013-11-16
     */
    public void parseWord(XWPFDocument document, Map<String, Object> map) throws Exception {
        parseWordSetValue((MyXWPFDocument) document, map);
    }

    private void parseWordSetValue(MyXWPFDocument doc, Map<String, Object> map) throws Exception {
        parseAllParagraphic(doc.getParagraphs(), map);
        parseHeaderAndFoot(doc, map);
        XWPFTable table;
        Iterator<XWPFTable> itTable = doc.getTablesIterator();
        while (itTable.hasNext()) {
            table = itTable.next();
            if (table.getText().indexOf("{{") != -1) {
                parseThisTable(table, map);
            }
        }

    }

    /**
     * Resolve headers and footers
     *
     * @param doc
     * @param map
     * @throws Exception
     */
    private void parseHeaderAndFoot(MyXWPFDocument doc, Map<String, Object> map) throws Exception {
        List<XWPFHeader> headerList = doc.getHeaderList();
        for (XWPFHeader xwpfHeader : headerList) {
            for (int i = 0; i < xwpfHeader.getListParagraph().size(); i++) {
                parseThisParagraph(xwpfHeader.getListParagraph().get(i), map);
            }
        }
        List<XWPFFooter> footerList = doc.getFooterList();
        for (XWPFFooter xwpfFooter : footerList) {
            for (int i = 0; i < xwpfFooter.getListParagraph().size(); i++) {
                parseThisParagraph(xwpfFooter.getListParagraph().get(i), map);
            }
        }

    }
}
