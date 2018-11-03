package com.poi.testpoi;

import com.poi.testpoi.common.MyException;
import com.poi.testpoi.pojo.User;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.junit.Before;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelTest2 {

    Workbook wb = null;

    @Before
    public void setup() throws Exception{
        InputStream is = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import.xls");
        wb = new HSSFWorkbook(is); //2003
    }

    @Test
    public void loadExcel5(){
        // 记录错误
        Map<String,Object> errors = new HashMap<>();

//        // 获取标题坐标信息
//        int rowIndex = 1;
//        int cellIndex = 1;
//        String titleExperssion = "";
//        boolean checkTitle = false;
//
//        // 获取标题
//        TitleModel titleModel = new TitleModel(rowIndex,cellIndex,titleExperssion,checkTitle);
//
//        getTitle(titleModel,errors); // return 返回值


//        int rowIndex = 2;
//        int cellStartIndex = 1;
//        int cellEndIndex = 2;
//        String[] cellNames = {"用户名称","用户密码"};
//        boolean checkHead = false;
//
//        HeadModel headModel = new HeadModel(rowIndex,cellStartIndex,cellEndIndex,cellNames,checkHead);
//
//        getHead(headModel,errors);

         // 获取数据
        int rowStartIndex = 3;
        int rowEndIndex = 11;
        int cellStartIndex = 1;
        int cellEndIndex = 2;
        Map<String,Object> columns = new HashMap<>();
        columns.put("username",1);
        columns.put("password",2);

        DataModel dataModel = new DataModel(rowStartIndex,rowEndIndex,cellStartIndex,cellEndIndex,columns);

        List<User> list = getData(dataModel,errors);
        for (User user : list){
            System.out.println(user.getUsername() + ","+user.getPassword());
        }

        System.out.println(errors);
    }

    // 获取数据
    public List<User> getData(DataModel dataModel,Map errors){
        Sheet sheet = wb.getSheetAt(0);

        int rowStartIndex = dataModel.getRowStartIndex();
        int rowEndIndex = dataModel.getRowEndIndex();
        int cellStartIndex = dataModel.getCellStartIndex();
        int cellEndIndex = dataModel.getCellEndIndex();
        Map<String,Object> columns = dataModel.getColumns();

        User user = null;
        List<User> list = new ArrayList<>();

        // 循环行数
        for (int i = rowStartIndex;i <= rowEndIndex;i++){
            Row r = sheet.getRow(i-1);

            if(r == null){
                errors.put("data","数据异常,坐标 row: "+i+",col:"+cellStartIndex);
            }

            int usernameIndex = (int) columns.get("username");
            int passwordIndex = (int) columns.get("password");

            user = new User();
            Cell username = r.getCell(usernameIndex-1);
            Cell password = r.getCell(passwordIndex-1);

            Object unameVal = getCellValue2(username.getCellType(),username);
            Object pwdVal = getCellValue2(password.getCellType(),password);

            if (unameVal==null){
                errors.put("data"+i+usernameIndex,"数据异常，坐标 row: "+i+",col:"+usernameIndex);
            }else if (pwdVal==null){
                errors.put("data"+i+passwordIndex,"数据异常，坐标 row: "+i+",col:"+passwordIndex);
            }

            if (errors.isEmpty()){
                user.setUsername(unameVal.toString());
                user.setPassword(pwdVal.toString());
                list.add(user);
            }

        }
        return list;
    }

    // 获取头信息
    public void getHead(HeadModel headModel,Map errors){
        Sheet sheet = wb.getSheetAt(0);

        int rowIndex = headModel.getRowIndex();
        int cellStartIndex = headModel.getCellStartIndex();
        int cellEndIndex = headModel.getCellEndIndex();
        String[] cellNames = headModel.getCellNames();
        boolean checkHead = headModel.isCheckHead();

        // 获取行数
        Row r = sheet.getRow(rowIndex-1);

        if(r == null){
            errors.put("title","标题为空，坐标 row: "+rowIndex+",col:"+cellStartIndex);
        }


        for (int i = cellStartIndex; i <= cellEndIndex; i++){
            Cell c = r.getCell(i-1);
            if (c == null) {
                errors.put("head"+rowIndex+i,"标题为空，坐标 row: "+rowIndex+",col:"+i);
            } else {
                CellType cellType = c.getCellType();
                Object xx = getCellValue2(cellType,c);
                System.out.println(xx);
            }
        }
    }

    // 获取标题
    public void getTitle(TitleModel titleModel, Map errors){
        Sheet sheet = wb.getSheetAt(0);

        int rowIndex = titleModel.getRowIndex();
        int cellIndex = titleModel.getCellIndex();
        // 获取行数
        Row r = sheet.getRow(rowIndex-1);

        if(r == null){
            errors.put("title","标题为空，坐标 row: "+rowIndex+",col:"+cellIndex);
        }

        // 根据行数获取列数
        Cell c = r.getCell(cellIndex-1);

        if (c == null) {
            errors.put("title","标题为空，坐标 row: "+rowIndex+",col:"+cellIndex);
        } else {
            CellType cellType = c.getCellType();
            Object xx = getCellValue2(cellType,c);
            System.out.println(xx);
        }
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

    @Test
    public void loadExcel4(){
        Sheet sheet = wb.getSheetAt(0);

        //确定要处理的行
        int rowStart = Math.min(3,sheet.getFirstRowNum());
        int rowEnd = Math.max(11,sheet.getLastRowNum());

        System.out.println(rowStart);
        System.out.println(rowEnd);


        for(int rowNum = rowStart; rowNum <rowEnd; rowNum ++){
            Row r = sheet.getRow(rowNum);
            if(r == null){
                //整行都是空的
                //根据需要处理它
                continue;
            }


            if (rowNum==(13-1)){
                System.out.println("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
            }

            int lastColumn = Math.max(r.getLastCellNum(), 1);
            System.out.println("last="+lastColumn);


            for (int cn = 0; cn < lastColumn; cn++) {
                Cell c = r.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (c == null) {
                    //此单元格中的电子表格为空
                } else {
                    //对单元格的内容做一些有用的事情
                    //System.out.println(c.getStringCellValue());
                    CellType cellType = c.getCellType();
                    getCellValue(cellType,c);
                }
            }
        }
    }

    public void getCellValue(CellType cellType,Cell c){
        if (cellType.equals(CellType.STRING)){
            System.out.println(c.getRichStringCellValue().getString());
        }

        if (cellType.equals(CellType.NUMERIC)){
            if (DateUtil.isCellDateFormatted(c)) {
                System.out.println(c.getDateCellValue());
            } else {
                System.out.println(c.getNumericCellValue());
            }
        }
        if (cellType.equals(CellType.BOOLEAN)){
            System.out.println(c.getBooleanCellValue());
        }

        if (cellType.equals(CellType.FORMULA)){
            System.out.println(c.getCellFormula());
        }

        if (cellType.equals(CellType.BLANK)){
            System.out.println("bank");
        }
    }

    @Test
    public void loadExce3() throws Exception{

        InputStream is = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import.xls");
        Workbook wb = new HSSFWorkbook(is); //2003

        DataFormatter formatter = new DataFormatter();
        Sheet sheet1 = wb.getSheetAt(0);
        for (Row row : sheet1) {

            System.out.println(row.getSheet().getCellComment(new CellAddress(new CellReference(3,1))));
            //在第一张纸上设置要从第4行到第5行重复的行。
            //CellRangeAddress cellAddresses = row.get().getRepeatingRows();
            //System.out.println(cellAddresses.getLastColumn());

            //（CellRangeAddress.valueOf（ “4:5”））;
            //将列设置为在第二张纸上从A列重复到C.
            //sheet2.setRepeatingColumns（CellRangeAddress.valueOf（ “A：C”））;


            for (Cell cell : row) {

                cell.getArrayFormulaRange();
//                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
//                System.out.print(cellRef.formatAsString());
//                System.out.print(" - ");
//
//                String text = formatter.formatCellValue(cell);
//                System.out.println(text);

            }
        }
    }



    @Test
    public void loadExcel2() throws Exception{
        InputStream is = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import.xls");
        Workbook wb = new HSSFWorkbook(is); //2003
        //Workbook wb = new XSSFWorkbook(is); //2007
        Sheet sheet = wb.getSheetAt(0);

        User user;
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {//r = 2 表示从第三行开始循环 如果你的第三行开始是数据
            Row row = sheet.getRow(r);//通过sheet表单对象得到 行对象
            if (row == null){
                continue;
            }


            System.out.println("=================总行数："+sheet.getLastRowNum());
            System.out.println("=================总列数："+row.getLastCellNum());
            //sheet.getLastRowNum() 的值是 10，所以Excel表中的数据至少是10条；不然报错 NullPointerException

            user = new User();
            String username = row.getCell(0).getStringCellValue();//得到每一行第一个单元格的值

            row.getCell(1).setCellType(CellType.STRING);//得到每一行的 第二个单元格的值
            String password = row.getCell(1).getStringCellValue();



            //完整的循环一次 就组成了一个对象
            user.setUsername(username);
            user.setPassword(password);
            System.out.println("username : "+username+", password : "+password);
        }
    }

    @Test
    public void loadExcel() throws Exception{
        InputStream is = new FileInputStream("D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import.xls");
        Workbook wb = new HSSFWorkbook(is); //2003
        //Workbook wb = new XSSFWorkbook(is); //2007
        Sheet sheet = wb.getSheetAt(0);

        User user;
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {//r = 2 表示从第三行开始循环 如果你的第三行开始是数据
            Row row = sheet.getRow(r);//通过sheet表单对象得到 行对象
            if (row == null){
                continue;
            }

            //sheet.getLastRowNum() 的值是 10，所以Excel表中的数据至少是10条；不然报错 NullPointerException

            user = new User();

            /*if( row.getCell(0).getCellType() !=1){//循环时，得到每一行的单元格进行判断
                throw new MyException("导入失败(第"+(r+1)+"行,用户名请设为文本格式)");
            }*/

            String username = row.getCell(0).getStringCellValue();//得到每一行第一个单元格的值

            if(username == null || username.isEmpty()){//判断是否为空
                throw new MyException("导入失败(第"+(r+1)+"行,用户名未填写)");
            }

//            row.getCell(1).setCellType(Cell.CELL_TYPE_STRING);//得到每一行的 第二个单元格的值
            row.getCell(1).setCellType(CellType.STRING);//得到每一行的 第二个单元格的值
            String password = row.getCell(1).getStringCellValue();


            if(password==null || password.isEmpty()){
                throw new MyException("导入失败(第"+(r+1)+"行,密码未填写)");
            }


            //完整的循环一次 就组成了一个对象
            user.setUsername(username);
            user.setPassword(password);
            System.out.println("username : "+username+", password : "+password);
        }
    }
}
