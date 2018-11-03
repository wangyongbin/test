package com.poi.testpoi.service.Impl;

import com.alibaba.fastjson.JSON;
import com.poi.testpoi.model.*;
import com.poi.testpoi.pojo.BiDeptCompleteDetail;
import com.poi.testpoi.pojo.BiDeptCostDetail;
import com.poi.testpoi.repository.BiDeptCompleteDetailRepository;
import com.poi.testpoi.repository.BiDeptCostDetailRepository;
import com.poi.testpoi.service.BiDeptService;
import com.poi.testpoi.util.ExcelImportUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Service
public class BiDeptServiceImpl implements BiDeptService {

	@Autowired
	private BiDeptCompleteDetailRepository biDeptCompleteDetailRepository;

	@Autowired
	private BiDeptCostDetailRepository biDeptCostDetailRepository;

	//@Transactional(readOnly = false,rollbackFor = Exception.class)
	@Override
	public boolean batchImport(String fileName, MultipartFile file) throws Exception {
		// 记录错误
		List<String> errors = new ArrayList<>();

		boolean notNull = false;

		boolean isExcel2003 = true;
		if (fileName.matches("^.+\\.(?i)(xlsx)$")) {
			isExcel2003 = false;
		}
		/*InputStream is = file.getInputStream();

		Workbook wb = null;
		if (isExcel2003) {
			wb = new HSSFWorkbook(is);
		} else {
			wb = new XSSFWorkbook(is);
		}*/

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
		//InputStream inputStream = new FileInputStream("import_dept.json");
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
				String titleVal = new ExcelImportUtils().getTitle(sheet,title,errors,sheetIndex-1);

				new ExcelImportUtils().getHead(sheet,head,errors,sheetIndex);

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

		return false;
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
				new ExcelImportUtils().setErrors(errors,i,cellStartIndex,sheetIndex);
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

			Object countMonthCellVal = new ExcelImportUtils().getCellValue2(countMonthCell.getCellType(),countMonthCell);
			Object yearTargetCellVal = new ExcelImportUtils().getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
			Object execValCellVal = new ExcelImportUtils().getCellValue2(execValCell.getCellType(),execValCell);
			Object execSumValCellVal = new ExcelImportUtils().getCellValue2(execSumValCell.getCellType(),execSumValCell);
			Object execSumRateCellVal = new ExcelImportUtils().getCellValue2(execSumRateCell.getCellType(),execSumRateCell);

			if (countMonthCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,countMonthIndex,sheetIndex);
			}else if (yearTargetCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,yearTargetIndex,sheetIndex);
			}else if (execValCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,execValIndex,sheetIndex);
			}else if (execSumValCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,execSumValIndex,sheetIndex);
			}else if (execSumRateCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,execSumRateIndex,sheetIndex);
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
				new ExcelImportUtils().setErrors(errors,i,cellStartIndex,sheetIndex);
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


			Object countMonthCellVal = new ExcelImportUtils().getCellValue2(countMonthCell.getCellType(),countMonthCell);
			Object yearTargetCellVal = new ExcelImportUtils().getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
			Object completeValCellVal = new ExcelImportUtils().getCellValue2(completeValCell.getCellType(),completeValCell);
			Object completeSumValCellVal = new ExcelImportUtils().getCellValue2(completeSumValCell.getCellType(),completeSumValCell);
			Object completeRateCellVal = new ExcelImportUtils().getCellValue2(completeRateCell.getCellType(),completeRateCell);
			Object lastYearValCellVal = new ExcelImportUtils().getCellValue2(lastYearValCell.getCellType(),lastYearValCell);
			Object lastYearIncreaseRateCellVal = new ExcelImportUtils().getCellValue2(lastYearIncreaseRateCell.getCellType(),lastYearIncreaseRateCell);

			if (countMonthCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,countMonthIndex,sheetIndex);
			}else if (yearTargetCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,yearTargetIndex,sheetIndex);
			}else if (completeValCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,completeValIndex,sheetIndex);
			}else if (completeSumValCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,completeSumValIndex,sheetIndex);
			}else if (completeRateCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,completeRateIndex,sheetIndex);
			}else if (lastYearValCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,lastYearValIndex,sheetIndex);
			}else if (lastYearIncreaseRateCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,lastYearIncreaseRateIndex,sheetIndex);
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

}
