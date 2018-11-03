package com.poi.testpoi;

public class HeadModel {
    private int rowIndex;// = 2;
    private int cellStartIndex;// = 1;
    private int cellEndIndex;// = 2;
    private String[] cellNames;// = {"用户名称","用户密码"};
    private boolean checkHead;// = false;


    public HeadModel() {
    }

    public HeadModel(int rowIndex, int cellStartIndex, int cellEndIndex, String[] cellNames, boolean checkHead) {
        this.rowIndex = rowIndex;
        this.cellStartIndex = cellStartIndex;
        this.cellEndIndex = cellEndIndex;
        this.cellNames = cellNames;
        this.checkHead = checkHead;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getCellStartIndex() {
        return cellStartIndex;
    }

    public void setCellStartIndex(int cellStartIndex) {
        this.cellStartIndex = cellStartIndex;
    }

    public int getCellEndIndex() {
        return cellEndIndex;
    }

    public void setCellEndIndex(int cellEndIndex) {
        this.cellEndIndex = cellEndIndex;
    }

    public String[] getCellNames() {
        return cellNames;
    }

    public void setCellNames(String[] cellNames) {
        this.cellNames = cellNames;
    }

    public boolean isCheckHead() {
        return checkHead;
    }

    public void setCheckHead(boolean checkHead) {
        this.checkHead = checkHead;
    }
}
