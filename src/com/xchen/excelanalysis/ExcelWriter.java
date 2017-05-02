package com.xchen.excelanalysis;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.xchen.excelanalysis.bean.DataBean;

public class ExcelWriter {

	public static void writer(String path, String fileName, List<DataBean> list, String title, String titleRow[])
			throws IOException {
		String fileType = path.substring(path.lastIndexOf(".") + 1);

		path = path.substring(0, path.lastIndexOf("\\"));

		Workbook wb = null;
		String excelPath = path + File.separator + fileName + "." + fileType;
		File file = new File(excelPath);
		Sheet sheet = null;
		// 创建工作文档对象
//		if (!file.exists()) {
//			if (fileType.equals("xls")) {
//				wb = new HSSFWorkbook();
//
//			} else if (fileType.equals("xlsx")) {
//
//				wb = new XSSFWorkbook();
//			} else {
//			}
//			// 创建sheet对象
//			sheet = (Sheet) wb.createSheet("结果");
//			OutputStream outputStream = new FileOutputStream(excelPath);
//			wb.write(outputStream);
//			outputStream.flush();
//			outputStream.close();
//
//		} else {
//			
//		}
		if (fileType.equals("xls")) {
			wb = new HSSFWorkbook();

		} else if (fileType.equals("xlsx")) {
			wb = new XSSFWorkbook();

		}
//		// 创建sheet对象
//		if (sheet == null) {
//			sheet = (Sheet) wb.createSheet("sheet1");
//			sheet.setDefaultColumnWidth((short) 15);
//		}
//
//		// 添加表头
//		Row row = sheet.createRow(0);
//		Cell cell = row.createCell(0);
//		row.setHeight((short) 540);
//		cell.setCellValue(title); // 创建第一行
//
//		CellStyle style = wb.createCellStyle(); // 样式对象
//		// 设置单元格的背景颜色为淡蓝色
//		style.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
//
//		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直
//		style.setAlignment(CellStyle.ALIGN_CENTER);// 水平
//		style.setWrapText(true);// 指定当单元格内容显示不下时自动换行
//
//		cell.setCellStyle(style); // 样式，居中
//
//		Font font = wb.createFont();
//		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
//		font.setFontName("宋体");
//		font.setFontHeight((short) 280);
//		style.setFont(font);
//		// 单元格合并
//		// 四个参数分别是：起始行，起始列，结束行，结束列
//		// sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));
//		// sheet.autoSizeColumn(5200);
//
//		row = sheet.createRow(1); // 创建第二行
//		// for (int i = 0; i < titleRow.length; i++) {
//		// cell = row.createCell(i);
//		// cell.setCellValue();
//		// cell.setCellStyle(style); // 样式，居中
//		// sheet.setColumnWidth(i, 20 * 256);
//		// }
//		//
//		for (short i = 0; i < titleRow.length; i++) {
//			cell = row.createCell(i);
//			cell.setCellStyle(style);
////			HSSFRichTextString text = new HSSFRichTextString();
//			cell.setCellValue(titleRow[i]);
//		}
//
//		row.setHeight((short) 540);

		
		  // 生成一个表格  
        sheet = wb.createSheet(title);  
        // 设置表格默认列宽度为15个字节  
        sheet.setDefaultColumnWidth((short) 15);  
        // 生成一个样式  
        CellStyle style = wb.createCellStyle();  
        // 设置这些样式  
        style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);  
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);  
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);  
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
        // 生成一个字体  
        Font font = wb.createFont();  
        font.setColor(HSSFColor.VIOLET.index);  
        font.setFontHeightInPoints((short) 12);  
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);  
        // 把字体应用到当前的样式  
        style.setFont(font);  
        // 生成并设置另一个样式  
        CellStyle style2 = wb.createCellStyle();  
        style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);  
        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);  
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);  
        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
        style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);  
        // 生成另一个字体  
        Font font2 = wb.createFont();  
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);  
        // 把字体应用到当前的样式  
        style2.setFont(font2);  
  
  
        // 产生表格标题行  
        Row row = sheet.createRow(0);  
        for (short i = 0; i < titleRow.length; i++)  
        {  
            Cell cell = row.createCell(i);  
            cell.setCellStyle(style);  
            cell.setCellValue(titleRow[i]);  
        }  
		// 循环写入行数据
		for (int i = 0; i < list.size(); i++) {
			if (list.get(i) == null) {
				continue;
			}
			row = (Row) sheet.createRow(i + 2);
			row.setHeight((short) 500);

			row.createCell(0).setCellValue((list.get(i)).name);
			row.createCell(1).setCellValue((list.get(i)).value);
			row.createCell(2).setCellValue((list.get(i)).nameDif);
			row.createCell(3).setCellValue((list.get(i)).valueDif);
			row.createCell(4).setCellValue((list.get(i)).log);

		}

		// 创建文件流
		OutputStream stream = new FileOutputStream(excelPath);
		// 写入数据
		wb.write(stream);
		// 关闭文件流
		stream.close();
	}

}
