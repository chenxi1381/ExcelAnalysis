package com.xchen.excelanalysis;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.xchen.excelanalysis.Window.ReadyToGo;
import com.xchen.excelanalysis.bean.DataBean;

public class Main {
	private Window w ;
	private List<DataBean> sameList = new ArrayList<DataBean>();
	private List<DataBean> unSameList = new ArrayList<DataBean>();
	public static void main(String[] args) {
		final Main main = new Main();
		main.w = new Window();
		main.w.setReadyToGoListener(new ReadyToGo() {
			@Override
			public void Ready(String path, String mSheetName1, String mSheetName2, String mName, String mValue) {
				List<DataBean> local1 = ExcelReader.getSheetData(path, mSheetName1, mName, mValue);
				List<DataBean> local2 = ExcelReader.getSheetData(path, mSheetName2, mName, mValue);
				main.analysis(local1,local2);
			
				
				try {
					String title[] = {mName,mValue};  
					ExcelWriter.writer(path,"same", main.sameList,"��ͬ������", title);
					String titleDif[] = {mSheetName1+" "+mName,mSheetName1+" "+mValue,mSheetName2+" "+mName,mSheetName2+" "+mValue};  
					ExcelWriter.writer(path,"different", main.unSameList,"����ͬ������", titleDif);
					main.w.setResult("�ɹ��ˡ�ԭĿ¼�����������ļ� same��different");
				}catch (IOException e) {
					// TODO Auto-generated catch block
					main.w.setResult("������"+e.getMessage());
					e.printStackTrace();
				}
			}
		});
	}

	

	private void analysis(List<DataBean> local1, List<DataBean> local2) {
	
		Levenshtein lt = new Levenshtein();

		for (DataBean bean1 : local1) {
			float maxSimilar = 0;
			DataBean similarBean = null;
			for (DataBean bean2 : local2) {
				float currentSimilar = 0;
				currentSimilar = lt.getSimilarityRatio(bean1.name, bean2.name);
				if (currentSimilar == 1) {
					similarBean = bean2;
					maxSimilar = 1;
					break;
				} else if (currentSimilar > maxSimilar) {
					similarBean = bean2;
					maxSimilar = currentSimilar;
				}
			}
			// TODO ��ʱ���԰�ƥ�䵽����ɾ����Ϊ��׼ȷ�ԣ��Ȳ�ɾ�� ����������
			if (similarBean == null) {
				DataBean dataBean = new DataBean();
				dataBean.name = bean1.name;
				dataBean.log = "û���ҵ����ʵ�ƥ����Ϣ";
				unSameList.add(dataBean);
				continue;
			}
			if (similarBean.value.equals(bean1.value)) {
				sameList.add(bean1);
				printSameLog(bean1);
			} else {
				DataBean dataBean = new DataBean();
				dataBean.name = bean1.name;
				dataBean.value = bean1.value;
				
				dataBean.nameDif = similarBean.name;
				dataBean.valueDif = similarBean.value;
				
				// TODO unSameLog ...
				unSameList.add(dataBean);
				printunSameLog(dataBean);
			}
		}

		System.out.println("done");

	}
	
	private static void printunSameLog(DataBean bean) {
		System.out.print("unsamelog   ���ƣ�" + bean.name + "       ֵ��" + bean.value);
		System.out.println("   ���ƣ�" + bean.nameDif + " ֵ��" + bean.valueDif);
	}

	private static void printSameLog(DataBean bean1) {
		System.out.println("samelog   ���ƣ�" + bean1.name + "         ֵ��" + bean1.value);
	}

	public static class Levenshtein {

		private int compare(String str, String target) {
			int d[][]; // ����
			int n = str.length();
			int m = target.length();
			int i; // ����str��
			int j; // ����target��
			char ch1; // str��
			char ch2; // target��
			int temp; // ��¼��ͬ�ַ�,��ĳ������λ��ֵ������,����0����1
			if (n == 0) {
				return m;
			}
			if (m == 0) {
				return n;
			}
			d = new int[n + 1][m + 1];
			for (i = 0; i <= n; i++) { // ��ʼ����һ��
				d[i][0] = i;
			}
			for (j = 0; j <= m; j++) { // ��ʼ����һ��
				d[0][j] = j;
			}
			for (i = 1; i <= n; i++) { // ����str
				ch1 = str.charAt(i - 1);
				// ȥƥ��target
				for (j = 1; j <= m; j++) {
					ch2 = target.charAt(j - 1);
					if (ch1 == ch2) {
						temp = 0;
					} else {
						temp = 1;
					}
					// ���+1,�ϱ�+1, ���Ͻ�+tempȡ��С
					d[i][j] = min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + temp);
				}
			}
			return d[n][m];
		}

		private int min(int one, int two, int three) {
			return (one = one < two ? one : two) < three ? one : three;
		}

		/**
		 * ��ȡ���ַ��������ƶ�
		 * 
		 * @param str
		 * @param target
		 * @return
		 */
		public float getSimilarityRatio(String str, String target) {
			return 1 - (float) compare(str, target) / Math.max(str.length(), target.length());
		}
	}
}
