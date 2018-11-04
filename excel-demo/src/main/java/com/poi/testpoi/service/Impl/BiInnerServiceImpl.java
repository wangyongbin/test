package com.poi.testpoi.service.Impl;

import com.alibaba.fastjson.JSON;
import com.poi.testpoi.model.*;
import com.poi.testpoi.pojo.BiGroupComplete;
import com.poi.testpoi.pojo.BiInnerComplete;
import com.poi.testpoi.pojo.BiLocalComplete;
import com.poi.testpoi.repository.BiInnerCompleteRepository;
import com.poi.testpoi.repository.BiLocalCompleteRepository;
import com.poi.testpoi.service.BiInnerCompleteService;
import com.poi.testpoi.service.BiLocalCompleteService;
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
public class BiInnerServiceImpl implements BiInnerCompleteService {

	@Autowired
	private BiInnerCompleteRepository biInnerCompleteRepository;

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

		fileName = "D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import_inner.xls";
		// 加载excel文件
		InputStream is = new FileInputStream(fileName);
		Workbook wb = null;
		if (isExcel2003) {
			wb = new HSSFWorkbook(is);
		} else {
			wb = new XSSFWorkbook(is);
		}

		String jsonName = "D:\\IdeaProjects\\my\\test\\excel-demo\\src\\main\\resources\\import_inner.json";
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

				if (tableName.equalsIgnoreCase("biInnerComplete")){
					List<BiInnerComplete> biInnerCompletes = getLocalCompleteData(sheet,data,errors,sheetIndex);

					//
					if (!errors.isEmpty()){
						for (String str : errors){
							System.out.println(str);
						}
					}else{
						//TODO 更新数据库
						System.out.println("更新数据库");
						for (BiInnerComplete bdcd : biInnerCompletes){
							System.out.println(bdcd.getCountYear()+","+bdcd.getTypeName()+","+bdcd.getValueUnit());
						}
					}
				}
			}

		}

		return false;
	}

	private List<BiInnerComplete> getLocalCompleteData(Sheet sheet, DataModel dataModel, List<String> errors, int sheetIndex) {

		int rowStartIndex = dataModel.getRowStartIndex();
		int rowEndIndex = dataModel.getRowEndIndex();
		int cellStartIndex = dataModel.getCellStartIndex();
		int cellEndIndex = dataModel.getCellEndIndex();
		Map<String,Object> columns = dataModel.getColumns();

		BiInnerComplete biDeptCompleteDetail = null;
		List<BiInnerComplete> list = new ArrayList<>();

		// 循环行数
		for (int i = rowStartIndex;i <= rowEndIndex;i++){
			Row r = sheet.getRow(i-1);

			if(r == null){
				new ExcelImportUtils().setErrors(errors,i,cellStartIndex,sheetIndex);
			}

			// 获取Excel列与数据库列映射

			if (sheetIndex==4){
				String countYear = (String) columns.get("countYear");
				//String deptName = (String) columns.get("deptName");
				//String typeName = (String) columns.get("typeName");
				String countMonth = (String) columns.get("countMonth");
				String valueUnit = (String) columns.get("valueUnit");

				// 索引
				//int countMonthIndex = (int) columns.get("countMonth");
				int typeNameIndex = (int) columns.get("typeName");
				int yearTargetIndex = (int) columns.get("innerTarget");
				int completeValIndex = (int) columns.get("completeVal");
				int completeSumValIndex = (int) columns.get("completeSumVal");
				int completeRateIndex = (int) columns.get("completeRate");
				int lastYearValIndex = (int) columns.get("lastYearVal");
				int lastYearIncreaseRateIndex = (int) columns.get("lastYearIncreaseRate");


				biDeptCompleteDetail = new BiInnerComplete();

				//Cell countMonthCell = r.getCell(countMonthIndex-1);
				Cell typeNameCell = r.getCell(typeNameIndex-1);
				Cell yearTargetCell = r.getCell(yearTargetIndex-1);
				Cell completeValCell = r.getCell(completeValIndex-1);
				Cell completeSumValCell = r.getCell(completeSumValIndex-1);
				Cell completeRateCell = r.getCell(completeRateIndex-1);
				Cell lastYearValCell = r.getCell(lastYearValIndex-1);
				Cell lastYearIncreaseRateCell = r.getCell(lastYearIncreaseRateIndex-1);


				//Object countMonthCellVal = new ExcelImportUtils().getCellValue2(countMonthCell.getCellType(),countMonthCell);
				Object typeNameCellVal = new ExcelImportUtils().getCellValue2(typeNameCell.getCellType(),typeNameCell);
				Object yearTargetCellVal = new ExcelImportUtils().getCellValue2(yearTargetCell.getCellType(),yearTargetCell);
				Object completeValCellVal = new ExcelImportUtils().getCellValue2(completeValCell.getCellType(),completeValCell);
				Object completeSumValCellVal = new ExcelImportUtils().getCellValue2(completeSumValCell.getCellType(),completeSumValCell);
				Object completeRateCellVal = new ExcelImportUtils().getCellValue2(completeRateCell.getCellType(),completeRateCell);
				Object lastYearValCellVal = new ExcelImportUtils().getCellValue2(lastYearValCell.getCellType(),lastYearValCell);
				Object lastYearIncreaseRateCellVal = new ExcelImportUtils().getCellValue2(lastYearIncreaseRateCell.getCellType(),lastYearIncreaseRateCell);

				if (typeNameCellVal==null){
					new ExcelImportUtils().setErrors(errors,i,typeNameIndex,sheetIndex);
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
					//biDeptCompleteDetail.setDeptName(deptName);
					//biDeptCompleteDetail.setTypeName(typeName);
					biDeptCompleteDetail.setCountMonth((String)countMonth);

					biDeptCompleteDetail.setValueUnit(valueUnit);

					//biDeptCompleteDetail.setCountMonth((String)countMonthCellVal);
					biDeptCompleteDetail.setTypeName((String)typeNameCellVal);
					//biDeptCompleteDetail.setYearTarget((Double) yearTargetCellVal);
					biDeptCompleteDetail.setInnerTarget((Double) yearTargetCellVal);
					biDeptCompleteDetail.setCompleteVal((Double)completeValCellVal);
					biDeptCompleteDetail.setCompleteSumVal((Double)completeSumValCellVal);
					biDeptCompleteDetail.setCompleteRate((Double)completeRateCellVal);
					biDeptCompleteDetail.setLastYearVal((Double)lastYearValCellVal);
					biDeptCompleteDetail.setLastYearIncreaseRate((Double)lastYearIncreaseRateCellVal);
					list.add(biDeptCompleteDetail);
				}
			}else {
				String countYear = (String) columns.get("countYear");
				//String deptName = (String) columns.get("deptName");
				String typeName = (String) columns.get("typeName");
				String valueUnit = (String) columns.get("valueUnit");

				// 索引
				int countMonthIndex = (int) columns.get("countMonth");
				int yearTargetIndex = (int) columns.get("innerTarget");
				int completeValIndex = (int) columns.get("completeVal");
				int completeSumValIndex = (int) columns.get("completeSumVal");
				int completeRateIndex = (int) columns.get("completeRate");
				int lastYearValIndex = (int) columns.get("lastYearVal");
				int lastYearIncreaseRateIndex = (int) columns.get("lastYearIncreaseRate");


				biDeptCompleteDetail = new BiInnerComplete();

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
					//biDeptCompleteDetail.setDeptName(deptName);
					biDeptCompleteDetail.setTypeName(typeName);
					biDeptCompleteDetail.setValueUnit(valueUnit);

					biDeptCompleteDetail.setCountMonth((String)countMonthCellVal);
					//biDeptCompleteDetail.setYearTarget((Double) yearTargetCellVal);
					biDeptCompleteDetail.setInnerTarget((Double) yearTargetCellVal);
					biDeptCompleteDetail.setCompleteVal((Double)completeValCellVal);
					biDeptCompleteDetail.setCompleteSumVal((Double)completeSumValCellVal);
					biDeptCompleteDetail.setCompleteRate((Double)completeRateCellVal);
					biDeptCompleteDetail.setLastYearVal((Double)lastYearValCellVal);
					biDeptCompleteDetail.setLastYearIncreaseRate((Double)lastYearIncreaseRateCellVal);
					list.add(biDeptCompleteDetail);
				}
			}

		}
		return list;

	}

}
