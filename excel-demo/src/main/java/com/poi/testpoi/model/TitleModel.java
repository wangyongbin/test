package com.poi.testpoi.model;

public class TitleModel {

    private int rowIndex;// = 1;
    private int cellIndex;// = 1;
    private String titleExperssion; // = "";
    private boolean checkTitle; // = true;

    public TitleModel() {
    }

    public TitleModel(int rowIndex, int cellIndex, String titleExperssion, boolean checkTitle) {
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
        this.titleExperssion = titleExperssion;
        this.checkTitle = checkTitle;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }

    public String getTitleExperssion() {
        return titleExperssion;
    }

    public void setTitleExperssion(String titleExperssion) {
        this.titleExperssion = titleExperssion;
    }

    public boolean isCheckTitle() {
        return checkTitle;
    }

    public void setCheckTitle(boolean checkTitle) {
        this.checkTitle = checkTitle;
    }
}
