package com.atviettelsolutions.easypoi.poi.common;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelUtil {

    public static void main(String[] args){
        //read excel data
        ArrayList<Map<String,String>> result = ExcelUtil.readExcelToObj("D:\\uploadForm.xlsx");
        for(Map<String,String> map:result){
            System.out.println(map);
        }

    }
    /**
     * read excel data
     * @param path
     */
    public static ArrayList<Map<String,String>> readExcelToObj(String path) {

        Workbook wb = null;
        ArrayList<Map<String,String>> result = null;
        try {
            wb = WorkbookFactory.create(new File(path));
            result = readExcel(wb, 0, 2, 0);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * read excel file
     * @param wb
     * @param sheetIndex sheet page index start,from 0
     * @param startReadLine the line to start reading starting, from 0
     * @param tailLine remove the last line read
     */
    public static ArrayList<Map<String,String>> readExcel(Workbook wb,int sheetIndex, int startReadLine, int tailLine) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        Row row = null;
        ArrayList<Map<String,String>> result = new ArrayList<Map<String,String>>();
        for(int i=startReadLine; i<sheet.getLastRowNum()-tailLine+1; i++) {

            row = sheet.getRow(i);
            Map<String,String> map = new HashMap<String,String>();
            for(Cell c : row) {
                String returnStr = "";

                boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
                //check if you need merge cell
                if(isMerge) {
                    String rs = getMergedRegionValue(sheet, row.getRowNum(), c.getColumnIndex());
//                    System.out.print(rs + "------ ");
                    returnStr = rs;
                }else {
//                    System.out.print(c.getRichStringCellValue()+"++++ ");
                    returnStr = c.getRichStringCellValue().getString();
                }
                if(c.getColumnIndex()==0){
                    map.put("id",returnStr);
                }else if(c.getColumnIndex()==1){
                    map.put("base",returnStr);
                }else if(c.getColumnIndex()==2){
                    map.put("siteName",returnStr);
                }else if(c.getColumnIndex()==3){
                    map.put("articleName",returnStr);
                }else if(c.getColumnIndex()==4){
                    map.put("mediaName",returnStr);
                }else if(c.getColumnIndex()==5){
                    map.put("mediaUrl",returnStr);
                }else if(c.getColumnIndex()==6){
                    map.put("newsSource",returnStr);
                }else if(c.getColumnIndex()==7){
                    map.put("isRecord",returnStr);
                }else if(c.getColumnIndex()==8){
                    map.put("recordTime",returnStr);
                }else if(c.getColumnIndex()==9){
                    map.put("remark",returnStr);
                }

            }
            result.add(map);
//            System.out.println();

        }
        return result;

    }

    /**
     * Get the value of the merged cell
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static String getMergedRegionValue(Sheet sheet ,int row , int column){
        int sheetMergeCount = sheet.getNumMergedRegions();

        for(int i = 0 ; i < sheetMergeCount ; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if(row >= firstRow && row <= lastRow){

                if(column >= firstColumn && column <= lastColumn){
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell) ;
                }
            }
        }

        return null ;
    }

    /**
     * Judgment merged rows
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static boolean isMergedRow(Sheet sheet,int row ,int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row == firstRow && row == lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Determine whether the specified cell is a merged cell
     * @param sheet
     * @param row row subscript
     * @param column column subscript
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet,int row ,int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Determine whether the sheet contains merged cells
     * @param sheet
     * @return
     */
    public static boolean hasMerged(Sheet sheet) {
        return sheet.getNumMergedRegions() > 0 ? true : false;
    }

    /**
     * Merge Cells
     * @param sheet
     * @param firstRow start line
     * @param lastRow end line
     * @param firstCol start col
     * @param lastCol end col
     */
    public static void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        try{
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        }catch (IllegalArgumentException e){
            e.fillInStackTrace();
        }
    }


    public static String getCellValue(Cell cell){

        if(cell == null) {
            return "";
        }

        if(cell.getCellTypeEnum() == CellType.STRING){

            return cell.getStringCellValue();

        }else if(cell.getCellTypeEnum() == CellType.BOOLEAN){

            return String.valueOf(cell.getBooleanCellValue());

        }else if(cell.getCellTypeEnum() ==  CellType.FORMULA){

            return cell.getCellFormula() ;

        }else if(cell.getCellTypeEnum() == CellType.NUMERIC){

            return String.valueOf(cell.getNumericCellValue());

        }
        return "";
    }
    

    public static String remove0Suffix(Object value){
    	if(value!=null) {
			String val = value.toString();
			if(val.endsWith(".0") && isNumberString(val)) {
				val = val.replace(".0", "");
			}
			return val;
		}
        return null;
    }

     private static boolean isNumberString(String str){
        String regex = "^[0-9]+\\.0+$";
        Pattern pattern = Pattern.compile(regex);
        Matcher m = pattern.matcher(str);
        if (m.find()) {
            return true;
        }
        return false;
    }
    

    public static void readContent(String fileName)  {
        boolean isE2007 = false;    //Determine whether it is in excel 2007 format
        if(fileName.endsWith("xlsx"))
            isE2007 = true;
        try {
            InputStream input = new FileInputStream(fileName);  //Create an input stream
            //Initialize according to the file format (2003 or 2007)
            Workbook wb  = null;

            if(isE2007)
                wb = new XSSFWorkbook(input);
            else
                wb = new HSSFWorkbook(input);
            Sheet sheet = wb.getSheetAt(0);     //get the first form
            Iterator<Row> rows = sheet.rowIterator(); //get the iterator of the first form
            while (rows.hasNext()) {
                Row row = rows.next();  //get row data
                System.out.println("Row #" + row.getRowNum());  //GetLineNumbersStartingFrom0
                Iterator<Cell> cells = row.cellIterator();    //get the iterator for the first row
                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    System.out.println("Cell #" + cell.getColumnIndex());
                    switch (cell.getCellTypeEnum()) {   //Output data according to the type in the cell
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue());
                            break;
                        case STRING:
                            System.out.println(cell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            System.out.println(cell.getCellFormula());
                            break;
                        default:
                            System.out.println("unsuported sell type======="+cell.getCellTypeEnum());
                            break;
                    }
                }
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

}