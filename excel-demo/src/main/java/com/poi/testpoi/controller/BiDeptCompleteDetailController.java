package com.poi.testpoi.controller;

import com.poi.testpoi.common.MyException;
import com.poi.testpoi.service.BiDeptService;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpSession;

@Controller
public class BiDeptCompleteDetailController {

	@Autowired
	private BiDeptService biDeptService;

	@RequestMapping(value = "/import")
	public String exImport(@RequestParam(value = "filename")MultipartFile file, HttpSession session) {

		boolean a = false;

		String fileName = file.getOriginalFilename();

		if (!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$")) {
			throw new MyException("上传文件格式不正确");
		}

		try {
			a = biDeptService.batchImport(fileName, file);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return "数据导入成功";
	}



















	/**
	 * 获取样式
	 *
	 * @param hssfWorkbook
	 * @param styleNum
	 * @return
	 */
	public HSSFCellStyle getStyle(HSSFWorkbook hssfWorkbook, Integer styleNum) {
		HSSFCellStyle style = hssfWorkbook.createCellStyle();
		style.setBorderRight(BorderStyle.THIN);//右边框
		style.setBorderBottom(BorderStyle.THIN);//下边框

		HSSFFont font = hssfWorkbook.createFont();
		font.setFontName("微软雅黑");//设置字体为微软雅黑

		HSSFPalette palette = hssfWorkbook.getCustomPalette();//拿到palette颜色板,可以根据需要设置颜色
		switch (styleNum) {
			case (0): {//HorizontalAlignment
				style.setAlignment(HorizontalAlignment.CENTER_SELECTION);//跨列居中
				font.setBold(true);//粗体
				font.setFontHeightInPoints((short) 14);//字体大小
				style.setFont(font);
//				palette.setColorAtIndex(HSSFColor.BLUE.index, (byte) 184, (byte) 204, (byte) 228);//替换颜色板中的颜色
//				style.setFillForegroundColor(HSSFColor.BLUE.index);
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
			break;
			case (1): {
				font.setBold(true);//粗体
				font.setFontHeightInPoints((short) 11);//字体大小
				style.setFont(font);
			}
			break;
			case (2): {
				font.setFontHeightInPoints((short) 10);
				style.setFont(font);
			}
			break;
			case (3): {
				style.setFont(font);

//				palette.setColorAtIndex(HSSFColor.GREEN.index, (byte) 0, (byte) 32, (byte) 96);//替换颜色板中的颜色
//				style.setFillForegroundColor(HSSFColor.GREEN.index);
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
			break;
		}

		return style;
	}


}
