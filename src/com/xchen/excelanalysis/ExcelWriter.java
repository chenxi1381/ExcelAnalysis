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
		// ���������ĵ�����
//		if (!file.exists()) {
//			if (fileType.equals("xls")) {
//				wb = new HSSFWorkbook();
//
//			} else if (fileType.equals("xlsx")) {
//
//				wb = new XSSFWorkbook();
//			} else {
//			}
//			// ����sheet����
//			sheet = (Sheet) wb.createSheet("���");
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
//		// ����sheet����
//		if (sheet == null) {
//			sheet = (Sheet) wb.createSheet("sheet1");
//			sheet.setDefaultColumnWidth((short) 15);
//		}
//
//		// ��ӱ�ͷ
//		Row row = sheet.createRow(0);
//		Cell cell = row.createCell(0);
//		row.setHeight((short) 540);
//		cell.setCellValue(title); // ������һ��
//
//		CellStyle style = wb.createCellStyle(); // ��ʽ����
//		// ���õ�Ԫ��ı�����ɫΪ����ɫ
//		style.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
//
//		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// ��ֱ
//		style.setAlignment(CellStyle.ALIGN_CENTER);// ˮƽ
//		style.setWrapText(true);// ָ������Ԫ��������ʾ����ʱ�Զ�����
//
//		cell.setCellStyle(style); // ��ʽ������
//
//		Font font = wb.createFont();
//		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
//		font.setFontName("����");
//		font.setFontHeight((short) 280);
//		style.setFont(font);
//		// ��Ԫ��ϲ�
//		// �ĸ������ֱ��ǣ���ʼ�У���ʼ�У������У�������
//		// sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));
//		// sheet.autoSizeColumn(5200);
//
//		row = sheet.createRow(1); // �����ڶ���
//		// for (int i = 0; i < titleRow.length; i++) {
//		// cell = row.createCell(i);
//		// cell.setCellValue();
//		// cell.setCellStyle(style); // ��ʽ������
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

		
		  // ����һ�����  
        sheet = wb.createSheet(title);  
        // ���ñ��Ĭ���п��Ϊ15���ֽ�  
        sheet.setDefaultColumnWidth((short) 15);  
        // ����һ����ʽ  
        CellStyle style = wb.createCellStyle();  
        // ������Щ��ʽ  
        style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);  
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);  
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);  
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
        // ����һ������  
        Font font = wb.createFont();  
        font.setColor(HSSFColor.VIOLET.index);  
        font.setFontHeightInPoints((short) 12);  
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);  
        // ������Ӧ�õ���ǰ����ʽ  
        style.setFont(font);  
        // ���ɲ�������һ����ʽ  
        CellStyle style2 = wb.createCellStyle();  
        style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);  
        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);  
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);  
        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
        style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);  
        // ������һ������  
        Font font2 = wb.createFont();  
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);  
        // ������Ӧ�õ���ǰ����ʽ  
        style2.setFont(font2);  
  
  
        // ������������  
        Row row = sheet.createRow(0);  
        for (short i = 0; i < titleRow.length; i++)  
        {  
            Cell cell = row.createCell(i);  
            cell.setCellStyle(style);  
            cell.setCellValue(titleRow[i]);  
        }  
		// ѭ��д��������
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

		// �����ļ���
		OutputStream stream = new FileOutputStream(excelPath);
		// д������
		wb.write(stream);
		// �ر��ļ���
		stream.close();
	}

}
