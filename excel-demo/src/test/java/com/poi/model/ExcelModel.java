package com.poi.model;

import java.util.List;
import java.util.Map;

/**
 *
 */
public class ExcelModel {

    private Map<String,Object> reqParams;
    private List<SheetModel> sheets;


    public Map<String, Object> getReqParams() {
        return reqParams;
    }

    public void setReqParams(Map<String, Object> reqParams) {
        this.reqParams = reqParams;
    }

    public List<SheetModel> getSheets() {
        return sheets;
    }

    public void setSheets(List<SheetModel> sheets) {
        this.sheets = sheets;
    }
}
