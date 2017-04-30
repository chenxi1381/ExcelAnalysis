package com.xchen.excelanalysis;

import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetAdapter;
import java.awt.dnd.DropTargetDropEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;

public class Window extends JFrame {

	private JPanel contentPane;
	private JTextField titleText;
	private JTextField filePath;
	private JTextField valueText;
	private JTextField sheet1Text;
	private JTextField sheet2Text;
	private JLabel resultTxt ;
	
	private String mFileName;
	private String mSheetName1;
	private String mSheetName2;
	private String mName;
	private String mValue;

	private ReadyToGo mReadyToGo;
	public interface ReadyToGo{
		public void Ready(String mFileName,String mSheetName1,String mSheetName2,String mName,String mValue);
	}
	public void setReadyToGoListener(ReadyToGo readyToGo){
		mReadyToGo = readyToGo;
	}
	/**
	 * Launch the application.
	 */
	

	/**
	 * Create the frame.
	 */
	public Window() {
		
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 326, 237);
		setTitle("tools");
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		GridBagLayout gbl_contentPane = new GridBagLayout();
		gbl_contentPane.columnWidths = new int[]{0, 0, 0, 0};
		gbl_contentPane.rowHeights = new int[]{0, 0, 0, 0, 0, 0, 0, 0};
		gbl_contentPane.columnWeights = new double[]{0.0, 1.0, 0.0, Double.MIN_VALUE};
		gbl_contentPane.rowWeights = new double[]{0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, Double.MIN_VALUE};
		contentPane.setLayout(gbl_contentPane);
		
		JLabel label = new JLabel("\u6587\u4EF6\u7684\u7EDD\u5BF9\u8DEF\u5F84\uFF1A");
		GridBagConstraints gbc_label = new GridBagConstraints();
		gbc_label.insets = new Insets(0, 0, 5, 5);
		gbc_label.anchor = GridBagConstraints.EAST;
		gbc_label.gridx = 0;
		gbc_label.gridy = 0;
		contentPane.add(label, gbc_label);
		
		filePath = new JTextField();
		filePath.setText("\u76F4\u63A5\u628A\u6587\u4EF6\u62D6\u5230\u8FD9\u91CC\u5C31\u884C");
		filePath.setToolTipText("");
		GridBagConstraints gbc_filePath = new GridBagConstraints();
		gbc_filePath.gridwidth = 2;
		gbc_filePath.insets = new Insets(0, 0, 5, 0);
		gbc_filePath.fill = GridBagConstraints.HORIZONTAL;
		gbc_filePath.gridx = 1;
		gbc_filePath.gridy = 0;
		contentPane.add(filePath, gbc_filePath);
		filePath.setColumns(10);
		
		JLabel lblSheet = new JLabel("sheet1\u540D\u79F0\uFF1A");
		GridBagConstraints gbc_lblSheet = new GridBagConstraints();
		gbc_lblSheet.anchor = GridBagConstraints.EAST;
		gbc_lblSheet.insets = new Insets(0, 0, 5, 5);
		gbc_lblSheet.gridx = 0;
		gbc_lblSheet.gridy = 1;
		contentPane.add(lblSheet, gbc_lblSheet);
		
		sheet1Text = new JTextField();
		GridBagConstraints gbc_sheet1Text = new GridBagConstraints();
		gbc_sheet1Text.gridwidth = 2;
		gbc_sheet1Text.insets = new Insets(0, 0, 5, 0);
		gbc_sheet1Text.fill = GridBagConstraints.BOTH;
		gbc_sheet1Text.gridx = 1;
		gbc_sheet1Text.gridy = 1;
		contentPane.add(sheet1Text, gbc_sheet1Text);
		sheet1Text.setColumns(10);
		
		JLabel lblSheet_1 = new JLabel("sheet2\u540D\u79F0\uFF1A");
		GridBagConstraints gbc_lblSheet_1 = new GridBagConstraints();
		gbc_lblSheet_1.anchor = GridBagConstraints.EAST;
		gbc_lblSheet_1.insets = new Insets(0, 0, 5, 5);
		gbc_lblSheet_1.gridx = 0;
		gbc_lblSheet_1.gridy = 2;
		contentPane.add(lblSheet_1, gbc_lblSheet_1);
		
		sheet2Text = new JTextField();
		GridBagConstraints gbc_sheet2Text = new GridBagConstraints();
		gbc_sheet2Text.gridwidth = 2;
		gbc_sheet2Text.insets = new Insets(0, 0, 5, 0);
		gbc_sheet2Text.fill = GridBagConstraints.HORIZONTAL;
		gbc_sheet2Text.gridx = 1;
		gbc_sheet2Text.gridy = 2;
		contentPane.add(sheet2Text, gbc_sheet2Text);
		sheet2Text.setColumns(10);
		
		JLabel label_1 = new JLabel("\u6807\u9898\u540D\u79F0\uFF1A");
		GridBagConstraints gbc_label_1 = new GridBagConstraints();
		gbc_label_1.anchor = GridBagConstraints.EAST;
		gbc_label_1.insets = new Insets(0, 0, 5, 5);
		gbc_label_1.gridx = 0;
		gbc_label_1.gridy = 3;
		contentPane.add(label_1, gbc_label_1);
		
		titleText = new JTextField();
		GridBagConstraints gbc_titleText = new GridBagConstraints();
		gbc_titleText.gridwidth = 2;
		gbc_titleText.insets = new Insets(0, 0, 5, 0);
		gbc_titleText.fill = GridBagConstraints.HORIZONTAL;
		gbc_titleText.gridx = 1;
		gbc_titleText.gridy = 3;
		contentPane.add(titleText, gbc_titleText);
		titleText.setColumns(10);
		
		JLabel label_2 = new JLabel("\u6570\u636E\u540D\u79F0\uFF1A");
		label_2.setToolTipText("yeah \u4F60\u53D1\u73B0\u4E86\u5F69\u86CB");
		GridBagConstraints gbc_label_2 = new GridBagConstraints();
		gbc_label_2.anchor = GridBagConstraints.EAST;
		gbc_label_2.insets = new Insets(0, 0, 5, 5);
		gbc_label_2.gridx = 0;
		gbc_label_2.gridy = 4;
		contentPane.add(label_2, gbc_label_2);
		
		valueText = new JTextField();
		GridBagConstraints gbc_valueText = new GridBagConstraints();
		gbc_valueText.insets = new Insets(0, 0, 5, 0);
		gbc_valueText.gridwidth = 2;
		gbc_valueText.fill = GridBagConstraints.HORIZONTAL;
		gbc_valueText.gridx = 1;
		gbc_valueText.gridy = 4;
		contentPane.add(valueText, gbc_valueText);
		valueText.setColumns(10);
		
		JButton button = new JButton("\u8BD5\u8BD5\u770B");
		GridBagConstraints gbc_button = new GridBagConstraints();
		gbc_button.insets = new Insets(0, 0, 5, 0);
		gbc_button.gridwidth = 3;
		gbc_button.fill = GridBagConstraints.HORIZONTAL;
		gbc_button.gridx = 0;
		gbc_button.gridy = 5;
		button.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent arg0) {
				if(mReadyToGo!= null){
					mReadyToGo.Ready(filePath.getText(), sheet1Text.getText(), sheet2Text.getText(), titleText.getText(), valueText.getText());
				}
			}
		});
		contentPane.add(button, gbc_button);
		
		resultTxt = new JLabel("\u7ED3\u679C");
		resultTxt.setVerticalAlignment(SwingConstants.TOP);
		GridBagConstraints gbc_resultTxt = new GridBagConstraints();
		gbc_resultTxt.anchor = GridBagConstraints.WEST;
		gbc_resultTxt.gridwidth = 3;
		gbc_resultTxt.insets = new Insets(0, 0, 0, 5);
		gbc_resultTxt.gridx = 0;
		gbc_resultTxt.gridy = 6;
		contentPane.add(resultTxt, gbc_resultTxt);
		drag();
		
		
		
		filePath.setText("D:\\test1.xlsx");
		sheet1Text.setText("sheet1");
		sheet2Text.setText("sheet2");
		titleText.setText("名称");
		valueText.setText("价格");
		setVisible(true);
	}
	public void setResult(String resultMsg){
		resultTxt.setText(resultMsg);
	}
	  public void drag()//定义的拖拽方法
	    {
	        //panel表示要接受拖拽的控件
	        new DropTarget(filePath, DnDConstants.ACTION_COPY_OR_MOVE, new DropTargetAdapter()
	        {
	            @Override
	            public void drop(DropTargetDropEvent dtde)//重写适配器的drop方法
	            {
	                try
	                {
	                    if (dtde.isDataFlavorSupported(DataFlavor.javaFileListFlavor))//如果拖入的文件格式受支持
	                    {
	                        dtde.acceptDrop(DnDConstants.ACTION_COPY_OR_MOVE);//接收拖拽来的数据
	                        List<File> list =  (List<File>) (dtde.getTransferable().getTransferData(DataFlavor.javaFileListFlavor));
	                        File file=list.get(0);
	                        filePath.setText(file.getAbsolutePath());
//	                        for(File file:list)
//	                            temp+=file.getAbsolutePath()+";\n";
//	                        JOptionPane.showMessageDialog(null, temp);
	                        dtde.dropComplete(true);//指示拖拽操作已完成
	                    }
	                    else
	                    {
	                        dtde.rejectDrop();//否则拒绝拖拽来的数据
	                    }
	                }
	                catch (Exception e)
	                {
	                    e.printStackTrace();
	                }
	            }
	        });
	    }
}
