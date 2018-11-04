//package com.poi.testpoi;
//
//import com.alibaba.fastjson.JSON;
//import com.b510.excel.vo.Student;
//import com.poi.testpoi.model.*;
//import com.poi.testpoi.pojo.User;
//import org.apache.commons.io.IOUtils;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.*;
//import org.junit.Before;
//import org.junit.Test;
//
//import java.io.FileInputStream;
//import java.io.InputStream;
//import java.util.*;
//
//public class ExcelTest {
//
//
//    @Before
//    public void setup() throws Exception{
//        /*InputStream is = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import.xls");
//        wb = new HSSFWorkbook(is); //2003*/
//    }
//
//    @Test
//    public void loadJsonFile() throws Exception{
//       // 加载json文件
//
//        InputStream inputStream = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import.json");
//        String text = IOUtils.toString(inputStream,"utf8");
//        ExcelModel excelModel = JSON.parseObject(text,ExcelModel.class);
//        System.out.println(JSON.toJSON(excelModel));
//    }
//
//    @Test
//    public void readJsonFileToDB() throws Exception{
//        // 记录错误
//        List<String> errors = new ArrayList<>();
//
//        // 加载excel文件
//        InputStream is = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import.xls");
//        Workbook wb = new HSSFWorkbook(is);
//
//        // 加载json文件
//        InputStream inputStream = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import.json");
//        String text = IOUtils.toString(inputStream,"utf8");
//        ExcelModel excelModel = JSON.parseObject(text,ExcelModel.class);
//
//        List<SheetModel> sheets = excelModel.getSheets();
//
//        for (SheetModel sheetModel : sheets){
//
//            String sheetTitle = sheetModel.getSheetTitle();
//            boolean checkTitle = sheetModel.isCheckTitle();
//
//            if (checkTitle){
//                // TODO 验证标题失败包含A|B|C
//                System.out.println("验证标题");
//            }
//
//            // sheet 页的索引
//            int sheetIndex = sheetModel.getSheetIndex();
//            Sheet sheet = wb.getSheetAt(sheetIndex-1);
//
//            // 数据区
//            List<DataAreaModel> dataAreaModels = sheetModel.getDataArea();
//
//            for (DataAreaModel dataAreaModel : dataAreaModels){
//                String tableName = dataAreaModel.getTableName();
//                TitleModel title = dataAreaModel.getTitle();
//                HeadModel head = dataAreaModel.getHead();
//                DataModel data = dataAreaModel.getData();
//
//                //TODO 验证标题参数的合法性
//                String titleVal = getTitle(sheet,title,errors);
//                validTitleModel(titleVal,title,errors);
//
//                getHead(sheet,head,errors);
//
//                List<User> users = getData(sheet,data,errors);
//
//                if (!errors.isEmpty()){
//                    for (String str : errors){
//                        System.out.println(str);
//                    }
//                }else{
//                    //TODO 更新数据库
//                    System.out.println("更新数据库");
//                    for (User user : users){
//                        System.out.println(user.getUsername()+","+user.getPassword());
//                    }
//                }
//            }
//
//        }
//    }
//
//    private void validTitleModel(String titleValue,TitleModel title, List<String> errors) {
//        //根据标题截取
//    }
//
//    // 获取数据
//    public List<User> getData(Sheet sheet,DataModel dataModel,List errors){
//        int rowStartIndex = dataModel.getRowStartIndex();
//        int rowEndIndex = dataModel.getRowEndIndex();
//        int cellStartIndex = dataModel.getCellStartIndex();
//        int cellEndIndex = dataModel.getCellEndIndex();
//        Map<String,Object> columns = dataModel.getColumns();
//
//        User user = null;
//        List<User> list = new ArrayList<>();
//
//        // 循环行数
//        for (int i = rowStartIndex;i <= rowEndIndex;i++){
//            Row r = sheet.getRow(i-1);
//
//            if(r == null){
//                setErrors(errors,i,cellStartIndex);
//            }
//
//            // 获取Excel列与数据库列映射
//            int usernameIndex = (int) columns.get("username");
//            int passwordIndex = (int) columns.get("password");
//
//            user = new User();
//            Cell username = r.getCell(usernameIndex-1);
//            Cell password = r.getCell(passwordIndex-1);
//
//            Object unameVal = getCellValue2(username.getCellType(),username);
//            Object pwdVal = getCellValue2(password.getCellType(),password);
//
//            if (unameVal==null){
//                setErrors(errors,i,usernameIndex);
//            }else if (pwdVal==null){
//                setErrors(errors,i,passwordIndex);
//            }
//
//            if (errors.isEmpty()){
//                user.setUsername(unameVal.toString());
//                user.setPassword(pwdVal.toString());
//                list.add(user);
//            }
//        }
//        return list;
//    }
//
//    // 获取头信息
//    public void getHead(Sheet sheet,HeadModel headModel,List errors){
//
//        int rowIndex = headModel.getRowIndex();
//        int cellStartIndex = headModel.getCellStartIndex();
//        int cellEndIndex = headModel.getCellEndIndex();
//        String[] cellNames = headModel.getCellNames();
//        boolean checkHead = headModel.isCheckHead();
//
//        // 获取行数
//        Row r = sheet.getRow(rowIndex-1);
//
//        if(r == null){
//            setErrors(errors,rowIndex,cellStartIndex);
//        }
//
//        for (int i = cellStartIndex; i <= cellEndIndex; i++){
//            Cell c = r.getCell(i-1);
//            if (c == null) {
//                setErrors(errors,rowIndex,i);
//            } else {
//                CellType cellType = c.getCellType();
//                Object objVal = getCellValue2(cellType,c);
//
//                // 是否为空
//                if (objVal == null){
//                    setErrors(errors,rowIndex,i);
//                }
//                //System.out.println(objVal);
//            }
//        }
//    }
//
//    // 获取标题
//    public String getTitle(Sheet sheet,TitleModel titleModel, List errors){
//
//        int rowIndex = titleModel.getRowIndex();
//        int cellIndex = titleModel.getCellIndex();
//        String titleExperssion = titleModel.getTitleExperssion();
//        boolean checkTitle = titleModel.isCheckTitle();
//
//        // 获取行数
//        Row r = sheet.getRow(rowIndex-1);
//
//        if(r == null){
//            setErrors(errors,rowIndex,cellIndex);
//        }
//
//        // 根据行数获取列数
//        Cell titleName = r.getCell(cellIndex-1);
//
//        Object titleNameVal = getCellValue2(titleName.getCellType(),titleName);
//
//        if (titleNameVal==null){
//            setErrors(errors,rowIndex,cellIndex);
//            return null;
//        }else{
//            return titleNameVal.toString();
//        }
//    }
//
//    public void setErrors(List errors,int i,int j){
//        errors.add("数据异常，单元格坐标，行："+i+"，列："+j);
//    }
//
//    public Object getCellValue2(CellType cellType,Cell c){
//        if (cellType.equals(CellType.STRING)){
//            return  c.getRichStringCellValue().getString();
//        }else if (cellType.equals(CellType.NUMERIC)){
//            if (DateUtil.isCellDateFormatted(c)) {
//                return c.getDateCellValue();
//            } else {
//                return c.getNumericCellValue();
//            }
//        }else if (cellType.equals(CellType.BOOLEAN)){
//            return c.getBooleanCellValue();
//        }else if (cellType.equals(CellType.FORMULA)){
//            return c.getCellFormula();
//        }else if (cellType.equals(CellType.BLANK)){
//            return null;
//        }
//        return null;
//    }
//
//
//
//
//}
