package com.poi.model;

import com.poi.testpoi.DataModel;
import com.poi.testpoi.HeadModel;
import com.poi.testpoi.TitleModel;


public class DataAreaModel {

    private String tableName;
    private TitleModel title;
    private HeadModel head;
    private DataModel data;

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    public TitleModel getTitle() {
        return title;
    }

    public void setTitle(TitleModel title) {
        this.title = title;
    }

    public HeadModel getHead() {
        return head;
    }

    public void setHead(HeadModel head) {
        this.head = head;
    }

    public DataModel getData() {
        return data;
    }

    public void setData(DataModel data) {
        this.data = data;
    }
}
