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
					ExcelWriter.writer(path,"same", main.sameList,"相同的数据", title);
					String titleDif[] = {mSheetName1+" "+mName,mSheetName1+" "+mValue,mSheetName2+" "+mName,mSheetName2+" "+mValue};  
					ExcelWriter.writer(path,"different", main.unSameList,"不相同的数据", titleDif);
					main.w.setResult("成功了。原目录下生成了俩文件 same和different");
				}catch (IOException e) {
					// TODO Auto-generated catch block
					main.w.setResult("错误了"+e.getMessage());
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
			// TODO 此时可以把匹配到的先删掉，为了准确性，先不删， 不考虑性能
			if (similarBean == null) {
				DataBean dataBean = new DataBean();
				dataBean.name = bean1.name;
				dataBean.log = "没有找到合适的匹配信息";
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
		System.out.print("unsamelog   名称：" + bean.name + "       值：" + bean.value);
		System.out.println("   名称：" + bean.nameDif + " 值：" + bean.valueDif);
	}

	private static void printSameLog(DataBean bean1) {
		System.out.println("samelog   名称：" + bean1.name + "         值：" + bean1.value);
	}

	public static class Levenshtein {

		private int compare(String str, String target) {
			int d[][]; // 矩阵
			int n = str.length();
			int m = target.length();
			int i; // 遍历str的
			int j; // 遍历target的
			char ch1; // str的
			char ch2; // target的
			int temp; // 记录相同字符,在某个矩阵位置值的增量,不是0就是1
			if (n == 0) {
				return m;
			}
			if (m == 0) {
				return n;
			}
			d = new int[n + 1][m + 1];
			for (i = 0; i <= n; i++) { // 初始化第一列
				d[i][0] = i;
			}
			for (j = 0; j <= m; j++) { // 初始化第一行
				d[0][j] = j;
			}
			for (i = 1; i <= n; i++) { // 遍历str
				ch1 = str.charAt(i - 1);
				// 去匹配target
				for (j = 1; j <= m; j++) {
					ch2 = target.charAt(j - 1);
					if (ch1 == ch2) {
						temp = 0;
					} else {
						temp = 1;
					}
					// 左边+1,上边+1, 左上角+temp取最小
					d[i][j] = min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + temp);
				}
			}
			return d[n][m];
		}

		private int min(int one, int two, int three) {
			return (one = one < two ? one : two) < three ? one : three;
		}

		/**
		 * 获取两字符串的相似度
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
