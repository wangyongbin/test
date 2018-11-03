package com.poi.testpoi;

import com.alibaba.fastjson.JSON;
import com.poi.testpoi.common.MyException;
import com.poi.testpoi.model.*;
import com.poi.testpoi.pojo.BiDeptCompleteDetail;
import com.poi.testpoi.pojo.BiDeptCostDetail;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class DeptExcelTest {

    @Test
    public void testLoadJsonFile() throws Exception{
        // 加载部门Excel对应json配置文件
        InputStream inputStream = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import_dept.json");
        String text = IOUtils.toString(inputStream,"utf8");
        ExcelModel excelModel = JSON.parseObject(text,ExcelModel.class);
        System.out.println(JSON.toJSON(excelModel));
    }

    @Test
    public void testReadJsonFileToDB() throws Exception{
        // 记录错误
        List<String> errors = new ArrayList<>();

        String fileName = "import_dept.xlsx";

        if (!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$")) {
            throw new MyException("上传文件格式不正确");
        }
        boolean isExcel2003 = true;
        if (fileName.matches("^.+\\.(?i)(xlsx)$")) {
            isExcel2003 = false;
        }

        // 加载excel文件
        InputStream is = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import_dept.xlsx");
        Workbook wb = null;
        if (isExcel2003) {
            wb = new HSSFWorkbook(is);
        } else {
            wb = new XSSFWorkbook(is);
        }

        String jsonName = "import_dept.json";
        // 加载json文件
        InputStream inputStream = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import_dept.json");
        String text = IOUtils.toString(inputStream,"utf8");
        ExcelModel excelModel = JSON.parseObject(text,ExcelModel.class);

        List<SheetModel> sheets = excelModel.getSheets();

        for (SheetModel sheetModel : sheets){

            String sheetTitle = sheetModel.getSheetTitle();
            boolean checkTitle = sheetModel.isCheckTitle();

            if (checkTitle){
                // TODO 是否验证标题 | 不验证
                //System.out.println("验证标题");
            }

            // sheet 页的索引
            int sheetIndex = sheetModel.getSheetIndex();
            Sheet sheet = wb.getSheetAt(sheetIndex-1);

            // 数据区
            List<DataAreaModel> dataAreaModels = sheetModel.getDataArea();

            for (DataAreaModel dataAreaModel : dataAreaModels){
                String tableName = dataAreaModel.getTableName();
                TitleModel title = dataAreaModel.getTitle();
                HeadModel head = dataAreaModel.getHead();
                DataModel data = dataAreaModel.getData();

                //TODO 是否验证标题参数的合法性 |  不验证
                String titleVal = getTitle(sheet,title,errors,sheetIndex-1);

                getHead(sheet,head,errors,sheetIndex);

                if (tableName.equalsIgnoreCase("biDeptCompleteDetail")){
                    List<BiDeptCompleteDetail> biDeptCompleteDetails = getBiDeptCompleteDetailData(sheet,data,errors,sheetIndex);

                    //
                    if (!errors.isEmpty()){
                        for (String str : errors){
                            System.out.println(str);
                        }
                    }else{
                        //TODO 更新数据库
                        System.out.println("更新数据库");
                        for (BiDeptCompleteDetail bdcd : biDeptCompleteDetails){
                            System.out.println(bdcd.getCountYear()+","+bdcd.getDeptName()+","+bdcd.getYearTarget());
                        }
                    }
                } else if (tableName.equalsIgnoreCase("biDeptCostDetail")){
                    List<BiDeptCostDetail> biDeptCostDetails = getBiDeptCostDetailData(sheet,data,errors,sheetIndex);

                    //
                    if (!errors.isEmpty()){
                        for (String str : errors){
                            System.out.println(str);
                        }
                    }else{
                        //TODO 更新数据库
                        System.out.println("更新数据库");
                        for (BiDeptCostDetail bdcd : biDeptCostDetails){
                            System.out.println(bdcd.getCountYear()+","+bdcd.getDeptName()+","+bdcd.getYearTarget());
                        }
                    }
                }
            }

        }
    }

    private List<BiDeptCostDetail> getBiDeptCostDetailData(Sheet sheet, DataModel dataModel, List<String> errors, int sheetIndex) {

        int rowStartIndex = dataModel.getRowStartIndex();
        int rowEndIndex = dataModel.getRowEndIndex();
        int cellStartIndex = dataModel.getCellStartIndex();
        int cellEndIndex = dataModel.getCellEndIndex();
        Map<String,Object> columns = dataModel.getColumns();

        BiDeptCostDetail biDeptCostDetail = null;
        List<BiDeptCostDetail> list = new ArrayList<>();

        // 循环行数
        for (int i = rowStartIndex;i <= rowEndIndex;i++){
            Row r = sheet.getRow(i-1);

            if(r == null){
                setErrors(errors,i,cellStartIndex,sheetIndex);
            }

            // 获取Excel列与数据库列映射
            String countYear = (String) columns.get("countYear");
            String deptName = (String) columns.get("deptName");
            String typeName = (String) columns.get("typeName");
            String valueUnit = (String) columns.get("valueUnit");

            // 索引
            int countMonthIndex = (int) columns.get("countMonth");
            int yearTargetIndex = (int) columns.get("yearTarget");
            int execValIndex = (int) columns.get("execVal");
            int execSumValIndex = (int) columns.get("execSumVal");
            int execSumRateIndex = (int) columns.get("execSumRate");


            biDeptCostDetail = new BiDeptCostDetail();

            Cell countMonthCell = r.getCell(countMonthIndex-1);
            Cell yearTargetCell = r.getCell(yearTargetIndex-1);
            Cell execValCell = r.getCell(execValIndex-1);
            Cell execSumValCell = r.getCell(execSumValIndex-1);
            Cell execSumRateCell = r.getCell(execSumRateIndex-1);

            Object countMonthCellVal = getCellValue2(countMonthCell.getCellType(),countMonthCell);
            Object yearTargetCellVal = getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
            Object execValCellVal = getCellValue2(execValCell.getCellType(),execValCell);
            Object execSumValCellVal = getCellValue2(execSumValCell.getCellType(),execSumValCell);
            Object execSumRateCellVal = getCellValue2(execSumRateCell.getCellType(),execSumRateCell);

            if (countMonthCellVal==null){
                setErrors(errors,i,countMonthIndex,sheetIndex);
            }else if (yearTargetCellVal==null){
                setErrors(errors,i,yearTargetIndex,sheetIndex);
            }else if (execValCellVal==null){
                setErrors(errors,i,execValIndex,sheetIndex);
            }else if (execSumValCellVal==null){
                setErrors(errors,i,execSumValIndex,sheetIndex);
            }else if (execSumRateCellVal==null){
                setErrors(errors,i,execSumRateIndex,sheetIndex);
            }


            if (errors.isEmpty()){

                biDeptCostDetail.setCountYear(countYear);
                biDeptCostDetail.setDeptName(deptName);
                biDeptCostDetail.setTypeName(typeName);
                biDeptCostDetail.setValueUnit(valueUnit);

                biDeptCostDetail.setCountMonth((String)countMonthCellVal);
                biDeptCostDetail.setYearTarget((Double) yearTargetCellVal);
                biDeptCostDetail.setExecVal((Double)execValCellVal);
                biDeptCostDetail.setExecSumVal((Double)execSumValCellVal);
                biDeptCostDetail.setExecSumRate((Double)execSumRateCellVal);
                list.add(biDeptCostDetail);
            }
        }
        return list;

    }

    private List<BiDeptCompleteDetail> getBiDeptCompleteDetailData(Sheet sheet, DataModel dataModel, List<String> errors, int sheetIndex) {

        int rowStartIndex = dataModel.getRowStartIndex();
        int rowEndIndex = dataModel.getRowEndIndex();
        int cellStartIndex = dataModel.getCellStartIndex();
        int cellEndIndex = dataModel.getCellEndIndex();
        Map<String,Object> columns = dataModel.getColumns();

        BiDeptCompleteDetail biDeptCompleteDetail = null;
        List<BiDeptCompleteDetail> list = new ArrayList<>();

        // 循环行数
        for (int i = rowStartIndex;i <= rowEndIndex;i++){
            Row r = sheet.getRow(i-1);

            if(r == null){
                setErrors(errors,i,cellStartIndex,sheetIndex);
            }

            // 获取Excel列与数据库列映射
            String countYear = (String) columns.get("countYear");
            String deptName = (String) columns.get("deptName");
            String typeName = (String) columns.get("typeName");
            String valueUnit = (String) columns.get("valueUnit");

            // 索引
            int countMonthIndex = (int) columns.get("countMonth");
            int yearTargetIndex = (int) columns.get("yearTarget");
            int completeValIndex = (int) columns.get("completeVal");
            int completeSumValIndex = (int) columns.get("completeSumVal");
            int completeRateIndex = (int) columns.get("completeRate");
            int lastYearValIndex = (int) columns.get("lastYearVal");
            int lastYearIncreaseRateIndex = (int) columns.get("lastYearIncreaseRate");


            biDeptCompleteDetail = new BiDeptCompleteDetail();

            Cell countMonthCell = r.getCell(countMonthIndex-1);
            Cell yearTargetCell = r.getCell(yearTargetIndex-1);
            Cell completeValCell = r.getCell(completeValIndex-1);
            Cell completeSumValCell = r.getCell(completeSumValIndex-1);
            Cell completeRateCell = r.getCell(completeRateIndex-1);
            Cell lastYearValCell = r.getCell(lastYearValIndex-1);
            Cell lastYearIncreaseRateCell = r.getCell(lastYearIncreaseRateIndex-1);


            Object countMonthCellVal = getCellValue2(countMonthCell.getCellType(),countMonthCell);
            Object yearTargetCellVal = getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
            Object completeValCellVal = getCellValue2(completeValCell.getCellType(),completeValCell);
            Object completeSumValCellVal = getCellValue2(completeSumValCell.getCellType(),completeSumValCell);
            Object completeRateCellVal = getCellValue2(completeRateCell.getCellType(),completeRateCell);
            Object lastYearValCellVal = getCellValue2(lastYearValCell.getCellType(),lastYearValCell);
            Object lastYearIncreaseRateCellVal = getCellValue2(lastYearIncreaseRateCell.getCellType(),lastYearIncreaseRateCell);

            if (countMonthCellVal==null){
                setErrors(errors,i,countMonthIndex,sheetIndex);
            }else if (yearTargetCellVal==null){
                setErrors(errors,i,yearTargetIndex,sheetIndex);
            }else if (completeValCellVal==null){
                setErrors(errors,i,completeValIndex,sheetIndex);
            }else if (completeSumValCellVal==null){
                setErrors(errors,i,completeSumValIndex,sheetIndex);
            }else if (completeRateCellVal==null){
                setErrors(errors,i,completeRateIndex,sheetIndex);
            }else if (lastYearValCellVal==null){
                setErrors(errors,i,lastYearValIndex,sheetIndex);
            }else if (lastYearIncreaseRateCellVal==null){
                setErrors(errors,i,lastYearIncreaseRateIndex,sheetIndex);
            }


            if (errors.isEmpty()){

                biDeptCompleteDetail.setCountYear(countYear);
                biDeptCompleteDetail.setDeptName(deptName);
                biDeptCompleteDetail.setTypeName(typeName);
                biDeptCompleteDetail.setValueUnit(valueUnit);

                biDeptCompleteDetail.setCountMonth((String)countMonthCellVal);
                biDeptCompleteDetail.setYearTarget((Double) yearTargetCellVal);
                biDeptCompleteDetail.setCompleteVal((Double)completeValCellVal);
                biDeptCompleteDetail.setCompleteSumVal((Double)completeSumValCellVal);
                biDeptCompleteDetail.setCompleteRate((Double)completeRateCellVal);
                biDeptCompleteDetail.setLastYearVal((Double)lastYearValCellVal);
                biDeptCompleteDetail.setLastYearIncreaseRate((Double)lastYearIncreaseRateCellVal);
                list.add(biDeptCompleteDetail);
            }
        }
        return list;

    }

    // 获取头信息
    public void getHead(Sheet sheet,HeadModel headModel,List errors,int sheetIndex){

        int rowIndex = headModel.getRowIndex();
        int cellStartIndex = headModel.getCellStartIndex();
        int cellEndIndex = headModel.getCellEndIndex();
        String[] cellNames = headModel.getCellNames();
        boolean checkHead = headModel.isCheckHead();

        // 获取行数
        Row r = sheet.getRow(rowIndex-1);

        if(r == null){
            setErrors(errors,rowIndex,cellStartIndex,sheetIndex);
        }

        for (int i = cellStartIndex; i <= cellEndIndex; i++){
            Cell c = r.getCell(i-1);
            if (c == null) {
                setErrors(errors,rowIndex,i,sheetIndex);
            } else {
                CellType cellType = c.getCellType();
                Object objVal = getCellValue2(cellType,c);

                // 是否为空
                if (objVal == null){
                    setErrors(errors,rowIndex,i,sheetIndex);
                }
                //System.out.println(objVal);
            }
        }
    }

    // 获取标题
    public String getTitle(Sheet sheet,TitleModel titleModel, List errors,int sheetIndex){

        int rowIndex = titleModel.getRowIndex();
        int cellIndex = titleModel.getCellIndex();
        String titleExperssion = titleModel.getTitleExperssion();
        boolean checkTitle = titleModel.isCheckTitle();

        // 获取行数
        Row r = sheet.getRow(rowIndex-1);

        if(r == null){
            setErrors(errors,rowIndex,cellIndex,sheetIndex);
        }

        // 根据行数获取列数
        Cell titleName = r.getCell(cellIndex-1);

        Object titleNameVal = getCellValue2(titleName.getCellType(),titleName);

        if (titleNameVal==null){
            setErrors(errors,rowIndex,cellIndex,sheetIndex);
            return null;
        }else{
            return titleNameVal.toString();
        }
    }

    public void setErrors(List errors,int i,int j,int sheetIndex){
        errors.add("数据异常，单元格坐标，第"+sheetIndex+"页，第"+i+"行，第"+j+"列.");
    }

    public Object getCellValue2(CellType cellType,Cell c){
        if (cellType.equals(CellType.STRING)){
            return  c.getRichStringCellValue().getString();
        }else if (cellType.equals(CellType.NUMERIC)){
            if (DateUtil.isCellDateFormatted(c)) {
                return c.getDateCellValue();
            } else {
                return c.getNumericCellValue();
            }
        }else if (cellType.equals(CellType.BOOLEAN)){
            return c.getBooleanCellValue();
        }else if (cellType.equals(CellType.FORMULA)){
            return c.getCellFormula();
        }else if (cellType.equals(CellType.BLANK)){
            return null;
        }
        return null;
    }
}
