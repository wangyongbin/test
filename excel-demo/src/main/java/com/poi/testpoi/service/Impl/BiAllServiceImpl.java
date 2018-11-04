package com.poi.testpoi.service.Impl;

import com.alibaba.fastjson.JSON;
import com.poi.testpoi.model.*;
import com.poi.testpoi.pojo.*;
import com.poi.testpoi.repository.*;
import com.poi.testpoi.service.BiAllService;
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
public class BiAllServiceImpl implements BiAllService {

	@Autowired
	private BiIndustryCompleteRepository biIndustryCompleteRepository;

	@Autowired
	private BiOwnerCompleteRepository biOwnerCompleteRepository;

	@Autowired
	private BiDeptCompleteRepository biDeptCompleteRepository;

	@Autowired
	private BiDeptProfitRepository biDeptProfitRepository;

	@Autowired
	private BiDeptCostRepository biDeptCostRepository;



	@Transactional(readOnly = false,rollbackFor = Exception.class)
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

		fileName = "D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import_all.xlsx";
		// 加载excel文件
		InputStream is = new FileInputStream(fileName);
		Workbook wb = null;
		if (isExcel2003) {
			wb = new HSSFWorkbook(is);
		} else {
			wb = new XSSFWorkbook(is);
		}

		String jsonName = "D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import_all.json";
		// 加载json文件
		InputStream inputStream = new FileInputStream(jsonName);
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

				// biIndustryComplete biOwnerComplete biDeptComplete biDeptProfit biDeptCost biDomainComplete
				if (tableName.equalsIgnoreCase("biIndustryComplete")){
					List<BiIndustryComplete> results = getBiIndustryCompleteData(sheet,data,errors,sheetIndex);
					//
					if (!errors.isEmpty()){
						for (String str : errors){
							System.out.println(str);
						}
					}else{
						//TODO 更新数据库
						System.out.println("更新数据库");
						for (BiIndustryComplete result : results){
							System.out.println(result.getCountYear()+","+result.getDeptName()+","+result.getYearTarget()+","+result.getTypeName());
						}
					}
				} else if (tableName.equalsIgnoreCase("biOwnerComplete")){

					List<BiOwnerComplete> results = getBiOwnerCompleteData(sheet,data,errors,sheetIndex);
					//
					if (!errors.isEmpty()){
						for (String str : errors){
							System.out.println(str);
						}
					}else{
						//TODO 更新数据库
						System.out.println("更新数据库");
						for (BiOwnerComplete result : results){
							System.out.println(result.getCountYear()+","+result.getDeptName()+","+result.getYearTarget()+","+result.getTypeName());
						}
					}

				} else if (tableName.equalsIgnoreCase("biDeptComplete")){

					List<BiDeptComplete> results = getBiDeptCompleteData(sheet,data,errors,sheetIndex);
					//
					if (!errors.isEmpty()){
						for (String str : errors){
							System.out.println(str);
						}
					}else{
						//TODO 更新数据库
						System.out.println("更新数据库");
						for (BiDeptComplete result : results){
							System.out.println(result.getCountYear()+","+result.getDeptName()+","+result.getYearTarget()+","+result.getTypeName());
						}
					}

				} else if (tableName.equalsIgnoreCase("biDeptProfit")){

					List<BiDeptProfit> results = getBiDeptProfitData(sheet,data,errors,sheetIndex);
					//
					if (!errors.isEmpty()){
						for (String str : errors){
							System.out.println(str);
						}
					}else{
						//TODO 更新数据库
						System.out.println("更新数据库");
						for (BiDeptProfit result : results){
							System.out.println(result.getCountYear()+","+result.getDeptName()+","+result.getYearTarget()+","+result.getTypeName());
						}
					}

				} else if (tableName.equalsIgnoreCase("biDeptCost")){

					List<BiDeptCost> results = getBiDeptCostData(sheet,data,errors,sheetIndex);
					//
					if (!errors.isEmpty()){
						for (String str : errors){
							System.out.println(str);
						}
					}else{
						//TODO 更新数据库
						System.out.println("更新数据库");
						for (BiDeptCost result : results){
							System.out.println(result.getCountYear()+","+result.getDeptName()+","+result.getYearTarget()+","+result.getTypeName());
						}
					}

				} else if (tableName.equalsIgnoreCase("biDomainComplete")){
					List<BiDomainComplete> results = getBiDomainCompleteData(sheet,data,errors,sheetIndex);
					//
					if (!errors.isEmpty()){
						for (String str : errors){
							System.out.println(str);
						}
					}else{
						//TODO 更新数据库
						System.out.println("更新数据库");
						for (BiDomainComplete result : results){
							System.out.println(result.getCountYear()+","+result.getDeptName()+","+result.getYearTarget()+","+result.getTypeName());
						}
					}
				}
			}

		}

		return false;
	}

	//领域新签、收入、到款情况
	private List<BiDomainComplete> getBiDomainCompleteData(Sheet sheet, DataModel dataModel, List<String> errors, int sheetIndex) {

		int rowStartIndex = dataModel.getRowStartIndex();
		int rowEndIndex = dataModel.getRowEndIndex();
		int cellStartIndex = dataModel.getCellStartIndex();
		int cellEndIndex = dataModel.getCellEndIndex();
		Map<String,Object> columns = dataModel.getColumns();

		BiDomainComplete entity = null;
		List<BiDomainComplete> list = new ArrayList<>();

		// 循环行数
		for (int i = rowStartIndex;i <= rowEndIndex;i++){
			Row r = sheet.getRow(i-1);

			if(r == null){
				new ExcelImportUtils().setErrors(errors,i,cellStartIndex,sheetIndex);
			}

			//
			String countYear = (String) columns.get("countYear");
			String typeName = (String) columns.get("typeName");
			String valueUnit = (String) columns.get("valueUnit");

			//
			//String ownerName = (String) columns.get("ownerName");
			String startMonth = (String) columns.get("startMonth");
			String entMonth = (String) columns.get("entMonth");

			// 索引
			int deptNameIndex = (int) columns.get("deptName");
			int yearTargetIndex = (int) columns.get("yearTarget");
			int completeValIndex = (int) columns.get("completeVal");
			int completeSumValIndex = (int) columns.get("completeSumVal");
			int completeRateIndex = (int) columns.get("completeRate");
			int lastYearValIndex = (int) columns.get("lastYearVal");
			int lastYearIncreaseRateIndex = (int) columns.get("lastYearIncreaseRate");

			entity = new BiDomainComplete();

			Cell deptNameCell = r.getCell(deptNameIndex-1);
			Cell yearTargetCell = r.getCell(yearTargetIndex-1);
			Cell completeValCell = r.getCell(completeValIndex-1);
			Cell completeSumValCell = r.getCell(completeSumValIndex-1);
			Cell completeRateCell = r.getCell(completeRateIndex-1);
			Cell lastYearValCell = r.getCell(lastYearValIndex-1);
			Cell lastYearIncreaseRateCell = r.getCell(lastYearIncreaseRateIndex-1);


			Object deptNameCellVal = new ExcelImportUtils().getCellValue2(deptNameCell.getCellType(),deptNameCell);
			Object yearTargetCellVal = new ExcelImportUtils().getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
			Object completeValCellVal = new ExcelImportUtils().getCellValue2(completeValCell.getCellType(),completeValCell);
			Object completeSumValCellVal = new ExcelImportUtils().getCellValue2(completeSumValCell.getCellType(),completeSumValCell);
			Object completeRateCellVal = new ExcelImportUtils().getCellValue2(completeRateCell.getCellType(),completeRateCell);
			Object lastYearValCellVal = new ExcelImportUtils().getCellValue2(lastYearValCell.getCellType(),lastYearValCell);
			Object lastYearIncreaseRateCellVal = new ExcelImportUtils().getCellValue2(lastYearIncreaseRateCell.getCellType(),lastYearIncreaseRateCell);

			if (deptNameCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,deptNameIndex,sheetIndex);
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

				entity.setCountYear(countYear);
				entity.setTypeName(typeName);
				entity.setValueUnit(valueUnit);

				//
				//entity.setOwnerName(ownerName);
				entity.setStartMonth(startMonth);
				entity.setEndMonth(entMonth);

				entity.setDeptName((String) deptNameCellVal);
				entity.setYearTarget((Double) yearTargetCellVal);
				entity.setCompleteVal((Double)completeValCellVal);
				entity.setCompleteSumVal((Double)completeSumValCellVal);
				entity.setCompleteRate((Double)completeRateCellVal);
				entity.setLastYearVal((Double)lastYearValCellVal);
				entity.setLastYearIncreaseRate((Double)lastYearIncreaseRateCellVal);
				list.add(entity);
			}

		}
		return list;
	}

	//部门收入、到款、新签、利润完成及成本控制情况 3 | 成本
	private List<BiDeptCost> getBiDeptCostData(Sheet sheet, DataModel dataModel, List<String> errors, int sheetIndex) {

		int rowStartIndex = dataModel.getRowStartIndex();
		int rowEndIndex = dataModel.getRowEndIndex();
		int cellStartIndex = dataModel.getCellStartIndex();
		int cellEndIndex = dataModel.getCellEndIndex();
		Map<String,Object> columns = dataModel.getColumns();

		BiDeptCost entity = null;
		List<BiDeptCost> list = new ArrayList<>();

		// 循环行数
		for (int i = rowStartIndex;i <= rowEndIndex;i++){
			Row r = sheet.getRow(i-1);

			if(r == null){
				new ExcelImportUtils().setErrors(errors,i,cellStartIndex,sheetIndex);
			}

			//
			String countYear = (String) columns.get("countYear");
			String typeName = (String) columns.get("typeName");
			String valueUnit = (String) columns.get("valueUnit");

			//
			//String ownerName = (String) columns.get("ownerName");
			String startMonth = (String) columns.get("startMonth");
			String entMonth = (String) columns.get("entMonth");

			// 索引
			int deptNameIndex = (int) columns.get("deptName");
			int yearTargetIndex = (int) columns.get("yearTarget");
			int completeValIndex = (int) columns.get("completeVal");
			int completeSumValIndex = (int) columns.get("completeSumVal");
			int completeRateIndex = (int) columns.get("completeRate");
//			int lastYearValIndex = (int) columns.get("lastYearVal");
//			int lastYearIncreaseRateIndex = (int) columns.get("lastYearIncreaseRate");

			entity = new BiDeptCost();

			Cell deptNameCell = r.getCell(deptNameIndex-1);
			Cell yearTargetCell = r.getCell(yearTargetIndex-1);
			Cell completeValCell = r.getCell(completeValIndex-1);
			Cell completeSumValCell = r.getCell(completeSumValIndex-1);
			Cell completeRateCell = r.getCell(completeRateIndex-1);
			/*Cell lastYearValCell = r.getCell(lastYearValIndex-1);
			Cell lastYearIncreaseRateCell = r.getCell(lastYearIncreaseRateIndex-1);*/


			Object deptNameCellVal = new ExcelImportUtils().getCellValue2(deptNameCell.getCellType(),deptNameCell);
			Object yearTargetCellVal = new ExcelImportUtils().getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
			Object completeValCellVal = new ExcelImportUtils().getCellValue2(completeValCell.getCellType(),completeValCell);
			Object completeSumValCellVal = new ExcelImportUtils().getCellValue2(completeSumValCell.getCellType(),completeSumValCell);
			Object completeRateCellVal = new ExcelImportUtils().getCellValue2(completeRateCell.getCellType(),completeRateCell);

			if (deptNameCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,deptNameIndex,sheetIndex);
			}else if (yearTargetCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,yearTargetIndex,sheetIndex);
			}else if (completeValCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,completeValIndex,sheetIndex);
			}else if (completeSumValCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,completeSumValIndex,sheetIndex);
			}else if (completeRateCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,completeRateIndex,sheetIndex);
			}

			if (errors.isEmpty()){

				entity.setCountYear(countYear);
				entity.setTypeName(typeName);
				entity.setValueUnit(valueUnit);

				//
				//entity.setOwnerName(ownerName);
				entity.setStartMonth(startMonth);
				entity.setEndMonth(entMonth);

				entity.setDeptName((String)
						deptNameCellVal);
				entity.setYearTarget((Double) yearTargetCellVal);
				entity.setCompleteVal((Double)completeValCellVal);
				entity.setCompleteSumVal((Double)completeSumValCellVal);
				entity.setCompleteRate((Double)completeRateCellVal);
				//entity.setLastYearVal((Double)lastYearValCellVal);
				//entity.setLastYearIncreaseRate((Double)lastYearIncreaseRateCellVal);
				list.add(entity);
			}

		}
		return list;
	}

	//部门收入、到款、新签、利润完成及成本控制情况 2 | 利润
	private List<BiDeptProfit> getBiDeptProfitData(Sheet sheet, DataModel dataModel, List<String> errors, int sheetIndex) {

		int rowStartIndex = dataModel.getRowStartIndex();
		int rowEndIndex = dataModel.getRowEndIndex();
		int cellStartIndex = dataModel.getCellStartIndex();
		int cellEndIndex = dataModel.getCellEndIndex();
		Map<String,Object> columns = dataModel.getColumns();

		BiDeptProfit entity = null;
		List<BiDeptProfit> list = new ArrayList<>();

		// 循环行数
		for (int i = rowStartIndex;i <= rowEndIndex;i++){
			Row r = sheet.getRow(i-1);

			if(r == null){
				new ExcelImportUtils().setErrors(errors,i,cellStartIndex,sheetIndex);
			}

			//
			String countYear = (String) columns.get("countYear");
			String typeName = (String) columns.get("typeName");
			String valueUnit = (String) columns.get("valueUnit");

			//
			//String ownerName = (String) columns.get("ownerName");
			String startMonth = (String) columns.get("startMonth");
			String entMonth = (String) columns.get("entMonth");

			// 索引
			int deptNameIndex = (int) columns.get("deptName");
			int yearTargetIndex = (int) columns.get("yearTarget");
			int completeValIndex = (int) columns.get("completeVal");
			int completeSumValIndex = (int) columns.get("completeSumVal");
			int completeRateIndex = (int) columns.get("completeRate");
//			int lastYearValIndex = (int) columns.get("lastYearVal");
//			int lastYearIncreaseRateIndex = (int) columns.get("lastYearIncreaseRate");

			entity = new BiDeptProfit();

			Cell deptNameCell = r.getCell(deptNameIndex-1);
			Cell yearTargetCell = r.getCell(yearTargetIndex-1);
			Cell completeValCell = r.getCell(completeValIndex-1);
			Cell completeSumValCell = r.getCell(completeSumValIndex-1);
			Cell completeRateCell = r.getCell(completeRateIndex-1);
//			Cell lastYearValCell = r.getCell(lastYearValIndex-1);
//			Cell lastYearIncreaseRateCell = r.getCell(lastYearIncreaseRateIndex-1);


			Object deptNameCellVal = new ExcelImportUtils().getCellValue2(deptNameCell.getCellType(),deptNameCell);
			Object yearTargetCellVal = new ExcelImportUtils().getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
			Object completeValCellVal = new ExcelImportUtils().getCellValue2(completeValCell.getCellType(),completeValCell);
			Object completeSumValCellVal = new ExcelImportUtils().getCellValue2(completeSumValCell.getCellType(),completeSumValCell);
			Object completeRateCellVal = new ExcelImportUtils().getCellValue2(completeRateCell.getCellType(),completeRateCell);
//			Object lastYearValCellVal = new ExcelImportUtils().getCellValue2(lastYearValCell.getCellType(),lastYearValCell);
//			Object lastYearIncreaseRateCellVal = new ExcelImportUtils().getCellValue2(lastYearIncreaseRateCell.getCellType(),lastYearIncreaseRateCell);

			if (deptNameCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,deptNameIndex,sheetIndex);
			}else if (yearTargetCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,yearTargetIndex,sheetIndex);
			}else if (completeValCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,completeValIndex,sheetIndex);
			}else if (completeSumValCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,completeSumValIndex,sheetIndex);
			}else if (completeRateCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,completeRateIndex,sheetIndex);
			}

			if (errors.isEmpty()){

				entity.setCountYear(countYear);
				entity.setTypeName(typeName);
				entity.setValueUnit(valueUnit);

				//
				//entity.setOwnerName(ownerName);
				entity.setStartMonth(startMonth);
				entity.setEndMonth(entMonth);

				entity.setDeptName((String)
						deptNameCellVal);
				entity.setYearTarget((Double) yearTargetCellVal);
				entity.setCompleteVal((Double)completeValCellVal);
				entity.setCompleteSumVal((Double)completeSumValCellVal);
				entity.setCompleteRate((Double)completeRateCellVal);
				list.add(entity);
			}

		}
		return list;
	}

	//部门收入、到款、新签、利润完成及成本控制情况 1 |
	private List<BiDeptComplete> getBiDeptCompleteData(Sheet sheet, DataModel dataModel, List<String> errors, int sheetIndex) {

		int rowStartIndex = dataModel.getRowStartIndex();
		int rowEndIndex = dataModel.getRowEndIndex();
		int cellStartIndex = dataModel.getCellStartIndex();
		int cellEndIndex = dataModel.getCellEndIndex();
		Map<String,Object> columns = dataModel.getColumns();

		BiDeptComplete entity = null;
		List<BiDeptComplete> list = new ArrayList<>();

		// 循环行数
		for (int i = rowStartIndex;i <= rowEndIndex;i++){
			Row r = sheet.getRow(i-1);

			if(r == null){
				new ExcelImportUtils().setErrors(errors,i,cellStartIndex,sheetIndex);
			}

			//
			String countYear = (String) columns.get("countYear");
			String typeName = (String) columns.get("typeName");
			String valueUnit = (String) columns.get("valueUnit");

			//
			//String ownerName = (String) columns.get("ownerName");
			String startMonth = (String) columns.get("startMonth");
			String entMonth = (String) columns.get("entMonth");

			// 索引
			int deptNameIndex = (int) columns.get("deptName");
			int yearTargetIndex = (int) columns.get("yearTarget");
			int completeValIndex = (int) columns.get("completeVal");
			int completeSumValIndex = (int) columns.get("completeSumVal");
			int completeRateIndex = (int) columns.get("completeRate");
			int lastYearValIndex = (int) columns.get("lastYearVal");
			int lastYearIncreaseRateIndex = (int) columns.get("lastYearIncreaseRate");

			entity = new BiDeptComplete();

			Cell deptNameCell = r.getCell(deptNameIndex-1);
			Cell yearTargetCell = r.getCell(yearTargetIndex-1);
			Cell completeValCell = r.getCell(completeValIndex-1);
			Cell completeSumValCell = r.getCell(completeSumValIndex-1);
			Cell completeRateCell = r.getCell(completeRateIndex-1);
			Cell lastYearValCell = r.getCell(lastYearValIndex-1);
			Cell lastYearIncreaseRateCell = r.getCell(lastYearIncreaseRateIndex-1);


			Object deptNameCellVal = new ExcelImportUtils().getCellValue2(deptNameCell.getCellType(),deptNameCell);
			Object yearTargetCellVal = new ExcelImportUtils().getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
			Object completeValCellVal = new ExcelImportUtils().getCellValue2(completeValCell.getCellType(),completeValCell);
			Object completeSumValCellVal = new ExcelImportUtils().getCellValue2(completeSumValCell.getCellType(),completeSumValCell);
			Object completeRateCellVal = new ExcelImportUtils().getCellValue2(completeRateCell.getCellType(),completeRateCell);
			Object lastYearValCellVal = new ExcelImportUtils().getCellValue2(lastYearValCell.getCellType(),lastYearValCell);
			Object lastYearIncreaseRateCellVal = new ExcelImportUtils().getCellValue2(lastYearIncreaseRateCell.getCellType(),lastYearIncreaseRateCell);

			if (deptNameCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,deptNameIndex,sheetIndex);
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

				entity.setCountYear(countYear);
				entity.setTypeName(typeName);
				entity.setValueUnit(valueUnit);

				//
				//entity.setOwnerName(ownerName);
				entity.setStartMonth(startMonth);
				entity.setEndMonth(entMonth);

				entity.setDeptName((String)
						deptNameCellVal);
				entity.setYearTarget((Double) yearTargetCellVal);
				entity.setCompleteVal((Double)completeValCellVal);
				entity.setCompleteSumVal((Double)completeSumValCellVal);
				entity.setCompleteRate((Double)completeRateCellVal);
				entity.setLastYearVal((Double)lastYearValCellVal);
				entity.setLastYearIncreaseRate((Double)lastYearIncreaseRateCellVal);
				list.add(entity);
			}

		}
		return list;
	}

	//经营主体指标完成情况
	private List<BiOwnerComplete> getBiOwnerCompleteData(Sheet sheet, DataModel dataModel, List<String> errors, int sheetIndex) {

		int rowStartIndex = dataModel.getRowStartIndex();
		int rowEndIndex = dataModel.getRowEndIndex();
		int cellStartIndex = dataModel.getCellStartIndex();
		int cellEndIndex = dataModel.getCellEndIndex();
		Map<String,Object> columns = dataModel.getColumns();

		BiOwnerComplete entity = null;
		List<BiOwnerComplete> list = new ArrayList<>();

		// 循环行数
		for (int i = rowStartIndex;i <= rowEndIndex;i++){
			Row r = sheet.getRow(i-1);

			if(r == null){
				new ExcelImportUtils().setErrors(errors,i,cellStartIndex,sheetIndex);
			}

			//
			String countYear = (String) columns.get("countYear");
			String typeName = (String) columns.get("typeName");
			String valueUnit = (String) columns.get("valueUnit");

			//
			//String ownerName = (String) columns.get("ownerName");
			String startMonth = (String) columns.get("startMonth");
			String entMonth = (String) columns.get("entMonth");

			// 索引
			int deptNameIndex = (int) columns.get("deptName");
			int yearTargetIndex = (int) columns.get("yearTarget");
			int completeValIndex = (int) columns.get("completeVal");
			int completeSumValIndex = (int) columns.get("completeSumVal");
			int completeRateIndex = (int) columns.get("completeRate");
			int lastYearValIndex = (int) columns.get("lastYearVal");
			int lastYearIncreaseRateIndex = (int) columns.get("lastYearIncreaseRate");

			entity = new BiOwnerComplete();

			Cell deptNameCell = r.getCell(deptNameIndex-1);
			Cell yearTargetCell = r.getCell(yearTargetIndex-1);
			Cell completeValCell = r.getCell(completeValIndex-1);
			Cell completeSumValCell = r.getCell(completeSumValIndex-1);
			Cell completeRateCell = r.getCell(completeRateIndex-1);
			Cell lastYearValCell = r.getCell(lastYearValIndex-1);
			Cell lastYearIncreaseRateCell = r.getCell(lastYearIncreaseRateIndex-1);


			Object deptNameCellVal = new ExcelImportUtils().getCellValue2(deptNameCell.getCellType(),deptNameCell);
			Object yearTargetCellVal = new ExcelImportUtils().getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
			Object completeValCellVal = new ExcelImportUtils().getCellValue2(completeValCell.getCellType(),completeValCell);
			Object completeSumValCellVal = new ExcelImportUtils().getCellValue2(completeSumValCell.getCellType(),completeSumValCell);
			Object completeRateCellVal = new ExcelImportUtils().getCellValue2(completeRateCell.getCellType(),completeRateCell);
			Object lastYearValCellVal = new ExcelImportUtils().getCellValue2(lastYearValCell.getCellType(),lastYearValCell);
			Object lastYearIncreaseRateCellVal = new ExcelImportUtils().getCellValue2(lastYearIncreaseRateCell.getCellType(),lastYearIncreaseRateCell);

			if (deptNameCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,deptNameIndex,sheetIndex);
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

				entity.setCountYear(countYear);
				entity.setTypeName(typeName);
				entity.setValueUnit(valueUnit);

				//
				//entity.setOwnerName(ownerName);
				entity.setStartMonth(startMonth);
				entity.setEndMonth(entMonth);

				entity.setDeptName((String)
						deptNameCellVal);
				entity.setYearTarget((Double) yearTargetCellVal);
				entity.setCompleteVal((Double)completeValCellVal);
				entity.setCompleteSumVal((Double)completeSumValCellVal);
				entity.setCompleteRate((Double)completeRateCellVal);
				entity.setLastYearVal((Double)lastYearValCellVal);
				entity.setLastYearIncreaseRate((Double)lastYearIncreaseRateCellVal);
				list.add(entity);
			}

		}
		return list;
	}

	//业态收入
	private List<BiIndustryComplete> getBiIndustryCompleteData(Sheet sheet, DataModel dataModel, List<String> errors, int sheetIndex) {

		int rowStartIndex = dataModel.getRowStartIndex();
		int rowEndIndex = dataModel.getRowEndIndex();
		int cellStartIndex = dataModel.getCellStartIndex();
		int cellEndIndex = dataModel.getCellEndIndex();
		Map<String,Object> columns = dataModel.getColumns();

		BiIndustryComplete entity = null;
		List<BiIndustryComplete> list = new ArrayList<>();

		// 循环行数
		for (int i = rowStartIndex;i <= rowEndIndex;i++){
			Row r = sheet.getRow(i-1);

			if(r == null){
				new ExcelImportUtils().setErrors(errors,i,cellStartIndex,sheetIndex);
			}

			//
			String countYear = (String) columns.get("countYear");
			String typeName = (String) columns.get("typeName");
			String valueUnit = (String) columns.get("valueUnit");

			//
			String ownerName = (String) columns.get("ownerName");
			String startMonth = (String) columns.get("startMonth");
			String entMonth = (String) columns.get("entMonth");

			// 索引
			int deptNameIndex = (int) columns.get("deptName");
			int yearTargetIndex = (int) columns.get("yearTarget");
			int completeValIndex = (int) columns.get("completeVal");
			int completeSumValIndex = (int) columns.get("completeSumVal");
			int completeRateIndex = (int) columns.get("completeRate");
			int lastYearValIndex = (int) columns.get("lastYearVal");
			int lastYearIncreaseRateIndex = (int) columns.get("lastYearIncreaseRate");

			entity = new BiIndustryComplete();

			Cell deptNameCell = r.getCell(deptNameIndex-1);
			Cell yearTargetCell = r.getCell(yearTargetIndex-1);
			Cell completeValCell = r.getCell(completeValIndex-1);
			Cell completeSumValCell = r.getCell(completeSumValIndex-1);
			Cell completeRateCell = r.getCell(completeRateIndex-1);
			Cell lastYearValCell = r.getCell(lastYearValIndex-1);
			Cell lastYearIncreaseRateCell = r.getCell(lastYearIncreaseRateIndex-1);


			Object deptNameCellVal = new ExcelImportUtils().getCellValue2(deptNameCell.getCellType(),deptNameCell);
			Object yearTargetCellVal = new ExcelImportUtils().getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
			Object completeValCellVal = new ExcelImportUtils().getCellValue2(completeValCell.getCellType(),completeValCell);
			Object completeSumValCellVal = new ExcelImportUtils().getCellValue2(completeSumValCell.getCellType(),completeSumValCell);
			Object completeRateCellVal = new ExcelImportUtils().getCellValue2(completeRateCell.getCellType(),completeRateCell);
			Object lastYearValCellVal = new ExcelImportUtils().getCellValue2(lastYearValCell.getCellType(),lastYearValCell);
			Object lastYearIncreaseRateCellVal = new ExcelImportUtils().getCellValue2(lastYearIncreaseRateCell.getCellType(),lastYearIncreaseRateCell);

			if (deptNameCellVal==null){
				new ExcelImportUtils().setErrors(errors,i,deptNameIndex,sheetIndex);
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

				entity.setCountYear(countYear);
				entity.setTypeName(typeName);
				entity.setValueUnit(valueUnit);

				//
				entity.setOwnerName(ownerName);
				entity.setStartMonth(startMonth);
				entity.setEndMonth(entMonth);

				entity.setDeptName((String)
						deptNameCellVal);
				entity.setYearTarget((Double) yearTargetCellVal);
				entity.setCompleteVal((Double)completeValCellVal);
				entity.setCompleteSumVal((Double)completeSumValCellVal);
				entity.setCompleteRate((Double)completeRateCellVal);
				entity.setLastYearVal((Double)lastYearValCellVal);
				entity.setLastYearIncreaseRate((Double)lastYearIncreaseRateCellVal);
				list.add(entity);
			}

		}
		return list;
	}
}
