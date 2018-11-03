package com.poi.testpoi.util;

import com.poi.testpoi.model.HeadModel;
import com.poi.testpoi.model.TitleModel;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

public class ExcelImportUtils {


    // @描述：是否是2003的excel，返回true是2003
    public static boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }

    //@描述：是否是2007的excel，返回true是2007
    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

    /**
     * 验证EXCEL文件
     *
     * @param filePath
     * @return
     */
    public static boolean validateExcel(String filePath) {
        if (filePath == null || !(isExcel2003(filePath) || isExcel2007(filePath))) {
            return false;
        }
        return true;
    }

    // 获取头信息
    public void getHead(Sheet sheet, HeadModel headModel, List errors, int sheetIndex) {

        int rowIndex = headModel.getRowIndex();
        int cellStartIndex = headModel.getCellStartIndex();
        int cellEndIndex = headModel.getCellEndIndex();
        String[] cellNames = headModel.getCellNames();
        boolean checkHead = headModel.isCheckHead();

        // 获取行数
        Row r = sheet.getRow(rowIndex - 1);

        if (r == null) {
            setErrors(errors, rowIndex, cellStartIndex, sheetIndex);
        }

        for (int i = cellStartIndex; i <= cellEndIndex; i++) {
            Cell c = r.getCell(i - 1);
            if (c == null) {
                setErrors(errors, rowIndex, i, sheetIndex);
            } else {
                CellType cellType = c.getCellType();
                Object objVal = getCellValue2(cellType, c);

                // 是否为空
                if (objVal == null) {
                    setErrors(errors, rowIndex, i, sheetIndex);
                }
                //System.out.println(objVal);
            }
        }
    }

    // 获取标题
    public String getTitle(Sheet sheet, TitleModel titleModel, List errors, int sheetIndex) {

        int rowIndex = titleModel.getRowIndex();
        int cellIndex = titleModel.getCellIndex();
        String titleExperssion = titleModel.getTitleExperssion();
        boolean checkTitle = titleModel.isCheckTitle();

        // 获取行数
        Row r = sheet.getRow(rowIndex - 1);

        if (r == null) {
            setErrors(errors, rowIndex, cellIndex, sheetIndex);
        }

        // 根据行数获取列数
        Cell titleName = r.getCell(cellIndex - 1);

        Object titleNameVal = getCellValue2(titleName.getCellType(), titleName);

        if (titleNameVal == null) {
            setErrors(errors, rowIndex, cellIndex, sheetIndex);
            return null;
        } else {
            return titleNameVal.toString();
        }
    }

    public void setErrors(List errors, int i, int j, int sheetIndex) {
        errors.add("数据异常，单元格坐标，第" + sheetIndex + "页，第" + i + "行，第" + j + "列.");
    }

    public Object getCellValue2(CellType cellType, Cell c) {
        if (cellType.equals(CellType.STRING)) {
            return c.getRichStringCellValue().getString();
        } else if (cellType.equals(CellType.NUMERIC)) {
            if (DateUtil.isCellDateFormatted(c)) {
                return c.getDateCellValue();
            } else {
                return c.getNumericCellValue();
            }
        } else if (cellType.equals(CellType.BOOLEAN)) {
            return c.getBooleanCellValue();
        } else if (cellType.equals(CellType.FORMULA)) {
            return c.getCellFormula();
        } else if (cellType.equals(CellType.BLANK)) {
            return null;
        }
        return null;
    }

}
