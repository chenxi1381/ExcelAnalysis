package com.xchen.excelanalysis;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


  
/** 
 * ��excel��ȡ����/��excel��д�� excel�б�ͷ����ͷÿ�е����ݶ�Ӧʵ��������� 
 *  
 * @author nagsh 
 *  
 */  
public class ExcelManage {  
	
	public static class InsuraceExcelBean{
		String InsuraceUser;
		String BankCardId;
		String IdCard;
		String BuyTime;
		String InsEndTime;
		String InsStartTime;
		String Money;
		String Type;
		public String getBankCardId() {
			return BankCardId;
		}

		public void setBankCardId(String bankCardId) {
			BankCardId = bankCardId;
		}

		public String getIdCard() {
			return IdCard;
		}

		public void setIdCard(String idCard) {
			IdCard = idCard;
		}

		public String getBuyTime() {
			return BuyTime;
		}

		public void setBuyTime(String buyTime) {
			BuyTime = buyTime;
		}

		public String getInsEndTime() {
			return InsEndTime;
		}

		public void setInsEndTime(String insEndTime) {
			InsEndTime = insEndTime;
		}

		public String getInsStartTime() {
			return InsStartTime;
		}

		public void setInsStartTime(String insStartTime) {
			InsStartTime = insStartTime;
		}

		public String getMoney() {
			return Money;
		}

		public void setMoney(String money) {
			Money = money;
		}

		public String getType() {
			return Type;
		}

		public void setType(String type) {
			Type = type;
		}

		public String getInsuraceUser() {
			return InsuraceUser;
		}

		public void setInsuraceUser(String insuraceUser) {
			InsuraceUser = insuraceUser;
		}
	
	}
    public static void main(String[] args) throws IOException {  
        String path = "D:/";  
        String fileName = "��������Ա�嵥(����)04";  
        String fileType = "xlsx";  
        List<InsuraceExcelBean> list = new ArrayList<>();
        for(int i=0; i<6; i++){
            InsuraceExcelBean bean = new InsuraceExcelBean();  
            bean.setInsuraceUser("test"+i);
            bean.setBankCardId("4444444444"+i+","+"55544444444444"+i+","+"999999999999999"+i);
            bean.setIdCard("666666"+i);
            bean.setBuyTime("2016-05-06");
            bean.setInsEndTime("2016-05-07");
            bean.setInsStartTime("2017-05-06");
            bean.setMoney("20,000");
            bean.setType("���");
            
            list.add(bean);
        }
        String title[] = {"������������","���֤��","�˻�����","���п���","���ս��(Ԫ)","����ʱ��","������Чʱ��","����ʧЧʱ��"};  
//        createExcel("E:/��������Ա�嵥(����).xlsx","sheet1",fileType,title);  
        
        writer(path, fileName, fileType,list,title);  
    }  
    
    @SuppressWarnings("resource")
    public static void writer(String path, String fileName,String fileType,List<InsuraceExcelBean> list,String titleRow[]) throws IOException {  
        Workbook wb = null; 
        String excelPath = path+File.separator+fileName+"."+fileType;
        File file = new File(excelPath);
        Sheet sheet =null;
        //���������ĵ�����   
        if (!file.exists()) {
            if (fileType.equals("xls")) {
                wb = new HSSFWorkbook();
                
            } else if(fileType.equals("xlsx")) {
                
                    wb = new XSSFWorkbook();
            } else {
            }
            //����sheet����   
            sheet = (Sheet) wb.createSheet("sheet1");  
            OutputStream outputStream = new FileOutputStream(excelPath);
            wb.write(outputStream);
            outputStream.flush();
            outputStream.close();
            
        } else {
            if (fileType.equals("xls")) {  
                wb = new HSSFWorkbook();  
                
            } else if(fileType.equals("xlsx")) { 
                wb = new XSSFWorkbook();  
                
            } else {  
            }  
        }
         //����sheet����   
        if (sheet==null) {
            sheet = (Sheet) wb.createSheet("sheet1");  
        }
        
        //��ӱ�ͷ  
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        row.setHeight((short) 540); 
        cell.setCellValue("��������Ա�嵥");    //������һ��    
        
        CellStyle style = wb.createCellStyle(); // ��ʽ����      
        // ���õ�Ԫ��ı�����ɫΪ����ɫ  
        style.setFillForegroundColor(HSSFColor.PALE_BLUE.index); 
        
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// ��ֱ      
        style.setAlignment(CellStyle.ALIGN_CENTER);// ˮƽ   
        style.setWrapText(true);// ָ������Ԫ��������ʾ����ʱ�Զ�����
       
        cell.setCellStyle(style); // ��ʽ������
        
        Font font = wb.createFont();  
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);  
        font.setFontName("����");  
        font.setFontHeight((short) 280);  
        style.setFont(font);  
        // ��Ԫ��ϲ�      
        // �ĸ������ֱ��ǣ���ʼ�У���ʼ�У������У�������      
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));  
        sheet.autoSizeColumn(5200);
        
        row = sheet.createRow(1);    //�����ڶ���    
        for(int i = 0;i < titleRow.length;i++){  
            cell = row.createCell(i);  
            cell.setCellValue(titleRow[i]);  
            cell.setCellStyle(style); // ��ʽ������
            sheet.setColumnWidth(i, 20 * 256); 
        }  
        row.setHeight((short) 540); 

        //ѭ��д��������   
        for (int i = 0; i < list.size(); i++) {  
            row = (Row) sheet.createRow(i+2);  
            row.setHeight((short) 500); 
            row.createCell(0).setCellValue(( list.get(i)).getInsuraceUser());
            row.createCell(1).setCellValue(( list.get(i)).getIdCard());
            row.createCell(2).setCellValue(( list.get(i)).getType());
            row.createCell(3).setCellValue(( list.get(i)).getBankCardId());
            row.createCell(4).setCellValue(( list.get(i)).getMoney());
            row.createCell(5).setCellValue(( list.get(i)).getBuyTime());
            row.createCell(6).setCellValue(( list.get(i)).getInsStartTime());
            row.createCell(7).setCellValue(( list.get(i)).getInsEndTime());
        }  
        
        //�����ļ���   
        OutputStream stream = new FileOutputStream(excelPath);  
        //д������   
        wb.write(stream);  
        //�ر��ļ���   
        stream.close();  
    }  
    
}