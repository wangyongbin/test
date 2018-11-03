package com.poi.testpoi.model;

import java.util.List;

public class SheetModel {

    private String sheetTitle;
    private boolean checkTitle;
    private int sheetIndex;
    private List<DataAreaModel> dataArea;

    public String getSheetTitle() {
        return sheetTitle;
    }

    public void setSheetTitle(String sheetTitle) {
        this.sheetTitle = sheetTitle;
    }

    public boolean isCheckTitle() {
        return checkTitle;
    }

    public void setCheckTitle(boolean checkTitle) {
        this.checkTitle = checkTitle;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public List<DataAreaModel> getDataArea() {
        return dataArea;
    }

    public void setDataArea(List<DataAreaModel> dataArea) {
        this.dataArea = dataArea;
    }
}
