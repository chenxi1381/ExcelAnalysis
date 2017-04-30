package com.xchen.excelanalysis;

import org.apache.log4j.spi.LoggerFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.logging.Logger;

public class WriterExcelUtil {

	    public static void main(String[] args) {
	        String path = "E://demo.xlsx";
	        String name = "test";
	        List<String> titles =Lists.newArrayList();
	        titles.add("id");
	        titles.add("name");
	        titles.add("age");
	        titles.add("birthday");
	        titles.add("gender");
	        titles.add("date");
	        List<Map<String, Object>> values = Lists.newArrayList();
	        for (int i = 0; i < 10; i++) {
	            Map<String, Object> map = Maps.newHashMap();
	            map.put("id", i + 1D);
	            map.put("name", "test_" + i);
	            map.put("age", i * 1.5);
	            map.put("gender", "man");
	            map.put("birthday", new Date());
	            map.put("date",  Calendar.getInstance());
	            values.add(map);
	        }
	        System.out.println(writerExcel(path, name, titles, values));
	    }

	    /**
	     * ����д��Excel�ļ�
	     *
	     * @param path �ļ�·���������ļ�ȫ�������磺D://file//demo.xls
	     * @param name sheet����
	     * @param titles �б�����
	     * @param values ���ݼ��ϣ�keyΪ���⣬valueΪ����
	     * @return True\False
	     */
	    public static boolean writerExcel(String path, String name, List<String> titles, List<Map<String, Object>> values) {
	        String style = path.substring(path.lastIndexOf("."), path.length()).toUpperCase(); // ���ļ�·���л�ȡ�ļ�������
	        return generateWorkbook(path, name, style, titles, values);
	    }

	    /**
	     * ������д��ָ��path�µ�Excel�ļ���
	     *
	     * @param path   �ļ��洢·��
	     * @param name   sheet��
	     * @param style  Excel����
	     * @param titles ���⴮
	     * @param values ���ݼ�
	     * @return True\False
	     */
	    private static boolean generateWorkbook(String path, String name, String style, List<String> titles, List<Map<String, Object>> values) {
	        Workbook workbook;
	        if ("XLS".equals(style.toUpperCase())) {
	            workbook = new HSSFWorkbook();
	        } else {
	            workbook = new XSSFWorkbook();
	        }
	        // ����һ�����
	        Sheet sheet;
	        if (null == name || "".equals(name)) {
	            sheet = workbook.createSheet(); // name Ϊ����ʹ��Ĭ��ֵ
	        } else {
	            sheet = workbook.createSheet(name);
	        }
	        // ���ñ��Ĭ���п��Ϊ15���ֽ�
	        sheet.setDefaultColumnWidth((short) 15);
	        // ������ʽ
	        Map<String, CellStyle> styles = createStyles(workbook);
	        /*
	         * ����������
	         */
	        Row row = sheet.createRow(0);
	        // �洢������Excel�ļ��е����
	        Map<String, Integer> titleOrder = Maps.newHashMap();
	        for (int i = 0; i < titles.size(); i++) {
	            Cell cell = row.createCell(i);
	            cell.setCellStyle(styles.get("header"));
	            String title = titles.get(i);
	            cell.setCellValue(title);
	            titleOrder.put(title, i);
	        }
	        /*
	         * д������
	         */
	        Iterator<Map<String, Object>> iterator = values.iterator();
	        int index = 0; // �к�
	        while (iterator.hasNext()) {
	            index++; // ��ȥ�����У��ӵ�һ�п�ʼд
	            row = sheet.createRow(index);
	            Map<String, Object> value = iterator.next();
	            for (Map.Entry<String, Object> map : value.entrySet()) {
	                // ��ȡ����
	                String title = map.getKey();
	                // ����������ȡ���
	                int i = titleOrder.get(title);
	                // ��ָ����Ŵ�����cell
	                Cell cell = row.createCell(i);
	                // ����cell����ʽ
	                if (index % 2 == 1) {
	                    cell.setCellStyle(styles.get("cellA"));
	                } else {
	                    cell.setCellStyle(styles.get("cellB"));
	                }
	                // ��ȡ�е�ֵ
	                Object object = map.getValue();
	                // �ж�object������
	                SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	                if (object instanceof Double) {
	                    cell.setCellValue((Double) object);
	                } else if (object instanceof Date) {
	                    String time = simpleDateFormat.format((Date) object);
	                    cell.setCellValue(time);
	                } else if (object instanceof Calendar) {
	                    Calendar calendar = (Calendar) object;
	                    String time = simpleDateFormat.format(calendar.getTime());
	                    cell.setCellValue(time);
	                } else if (object instanceof Boolean) {
	                    cell.setCellValue((Boolean) object);
	                } else {
	                    cell.setCellValue(object.toString());
	                }
	            }
	        }
	        /*
	         * д�뵽�ļ���
	         */
	        boolean isCorrect = false;
	        try {
	            File file = new File(path);
	            OutputStream outputStream = new FileOutputStream(file);
	            workbook.write(outputStream);
	            outputStream.close();
	            isCorrect = true;
	        } catch (IOException e) {
	            isCorrect = false;
	        }
	        try {
	            workbook.close();
	        } catch (IOException e) {
	            isCorrect = false;
	        }
	        return isCorrect;
	    }

	    /**
	     * Create a library of cell styles
	     */
	    /**
	     * @param wb
	     * @return
	     */
	    private static Map<String, CellStyle> createStyles(Workbook wb) {
	        Map<String, CellStyle> styles = Maps.newHashMap();

	        // ������ʽ
	        CellStyle titleStyle = wb.createCellStyle();
	        titleStyle.setAlignment(HorizontalAlignment.CENTER); // ˮƽ����
	        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER); // ��ֱ����
	        titleStyle.setLocked(true); // ��ʽ����
	        titleStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	        Font titleFont = wb.createFont();
	        titleFont.setFontHeightInPoints((short) 16);
	        titleFont.setBold(true);
	        titleFont.setFontName("΢���ź�");
	        titleStyle.setFont(titleFont);
	        styles.put("title", titleStyle);

	        // �ļ�ͷ��ʽ
	        CellStyle headerStyle = wb.createCellStyle();
	        headerStyle.setAlignment(HorizontalAlignment.CENTER);
	        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
	        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex()); // ǰ��ɫ
	        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // ��ɫ��䷽ʽ
	        headerStyle.setWrapText(true);
	        headerStyle.setBorderRight(BorderStyle.THIN); // ���ñ߽�
	        headerStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
	        headerStyle.setBorderLeft(BorderStyle.THIN);
	        headerStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
	        headerStyle.setBorderTop(BorderStyle.THIN);
	        headerStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
	        headerStyle.setBorderBottom(BorderStyle.THIN);
	        headerStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
	        Font headerFont = wb.createFont();
	        headerFont.setFontHeightInPoints((short) 12);
	        headerFont.setColor(IndexedColors.WHITE.getIndex());
	        titleFont.setFontName("΢���ź�");
	        headerStyle.setFont(headerFont);
	        styles.put("header", headerStyle);

	        Font cellStyleFont = wb.createFont();
	        cellStyleFont.setFontHeightInPoints((short) 12);
	        cellStyleFont.setColor(IndexedColors.BLUE_GREY.getIndex());
	        cellStyleFont.setFontName("΢���ź�");
	        
	        // ������ʽA
	        CellStyle cellStyleA = wb.createCellStyle();
	        cellStyleA.setAlignment(HorizontalAlignment.CENTER); // ��������
	        cellStyleA.setVerticalAlignment(VerticalAlignment.CENTER);
	        cellStyleA.setWrapText(true);
	        cellStyleA.setBorderRight(BorderStyle.THIN);
	        cellStyleA.setRightBorderColor(IndexedColors.BLACK.getIndex());
	        cellStyleA.setBorderLeft(BorderStyle.THIN);
	        cellStyleA.setLeftBorderColor(IndexedColors.BLACK.getIndex());
	        cellStyleA.setBorderTop(BorderStyle.THIN);
	        cellStyleA.setTopBorderColor(IndexedColors.BLACK.getIndex());
	        cellStyleA.setBorderBottom(BorderStyle.THIN);
	        cellStyleA.setBottomBorderColor(IndexedColors.BLACK.getIndex());
	        cellStyleA.setFont(cellStyleFont);
	        styles.put("cellA", cellStyleA);

	        // ������ʽB:���ǰ��ɫΪǳ��ɫ
	        CellStyle cellStyleB = wb.createCellStyle();
	        cellStyleB.setAlignment(HorizontalAlignment.CENTER);
	        cellStyleB.setVerticalAlignment(VerticalAlignment.CENTER);
	        cellStyleB.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	        cellStyleB.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        cellStyleB.setWrapText(true);
	        cellStyleB.setBorderRight(BorderStyle.THIN);
	        cellStyleB.setRightBorderColor(IndexedColors.BLACK.getIndex());
	        cellStyleB.setBorderLeft(BorderStyle.THIN);
	        cellStyleB.setLeftBorderColor(IndexedColors.BLACK.getIndex());
	        cellStyleB.setBorderTop(BorderStyle.THIN);
	        cellStyleB.setTopBorderColor(IndexedColors.BLACK.getIndex());
	        cellStyleB.setBorderBottom(BorderStyle.THIN);
	        cellStyleB.setBottomBorderColor(IndexedColors.BLACK.getIndex());
	        cellStyleB.setFont(cellStyleFont);
	        styles.put("cellB", cellStyleB);

	        return styles;
	    }
}
