package com.poi.testpoi;

import java.util.Map;

public class DataModel {
    private int rowStartIndex;// = 3;
    private int rowEndIndex;// = 11;
    private int cellStartIndex;// = 1;
    private int cellEndIndex;// = 2;

    private Map<String,Object> columns;//

    public DataModel() {
    }

    public DataModel(int rowStartIndex,int rowEndIndex, int cellStartIndex, int cellEndIndex,Map<String,Object> columns) {
        this.rowStartIndex = rowStartIndex;
        this.rowEndIndex = rowEndIndex;
        this.cellStartIndex = cellStartIndex;
        this.cellEndIndex = cellEndIndex;
        this.columns = columns;
    }

    public int getRowStartIndex() {
        return rowStartIndex;
    }

    public void setRowStartIndex(int rowStartIndex) {
        this.rowStartIndex = rowStartIndex;
    }

    public int getRowEndIndex() {
        return rowEndIndex;
    }

    public void setRowEndIndex(int rowEndIndex) {
        this.rowEndIndex = rowEndIndex;
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

    public Map<String, Object> getColumns() {
        return columns;
    }

    public void setColumns(Map<String, Object> columns) {
        this.columns = columns;
    }
}
