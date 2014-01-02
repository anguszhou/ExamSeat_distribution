package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;
import javax.swing.SwingWorker;

import com.linuxense.javadbf.DBFField;
import com.linuxense.javadbf.DBFReader;
import com.linuxense.javadbf.DBFWriter;

public class TaskExportTea extends SwingWorker<List<Teacher>, Teacher>{

	String filepath;
	private int lineCnt = 0;
	private int size = 0;
	int year=0;
	private DefaultProgressHandle progressHandle = null;
	
	public TaskExportTea(String filepath){
		this.filepath = filepath;
	}
	public void setProgressHandle(DefaultProgressHandle progressHandle) {
		this.progressHandle = progressHandle;
	}
	@Override
	protected List<Teacher> doInBackground() throws Exception {
		// TODO Auto-generated method stub
		String filename = filepath+File.separator+year+"jkap(已分配考场).dbf";
		System.out.println(filename);
		OutputStream fos = null;
		
		try {
			// 定义DBF文件字段
			// 分别定义各个字段信息，setFieldName和setName作用相同,
			if(size == 0){
				size = progressHandle.teaField.size();
			}
			DBFField[] fields = new DBFField[size];
			for (int i = 0; i < size; i++) {
				fields[i] = new DBFField();
				fields[i].setName(progressHandle.teaField.get(i));
				fields[i].setDataType(DBFField.FIELD_TYPE_C);
			}
			
			fields[0].setFieldLength(30);
			fields[1].setFieldLength(15);
			fields[2].setFieldLength(50);
			fields[3].setFieldLength(20);
			fields[4].setFieldLength(20);
			fields[5].setFieldLength(30);
			fields[6].setFieldLength(20);
			fields[7].setFieldLength(5);
			fields[8].setFieldLength(5);
			fields[9].setFieldLength(20);
			
			//DBFWriter writer = new DBFWriter(new File(path));
			// 定义DBFWriter实例用来写DBF文件
			DBFWriter writer = new DBFWriter();
			writer.setCharactersetName("GBK");
			// 把字段信息写入DBFWriter实例，即定义表结构
			writer.setFields(fields);
			// 一条条的写入记录
			Object[] rowData ;
			for(Teacher tt : progressHandle.teaList){
				rowData = new Object[size];
				
				Field[] teaField = tt.getClass().getDeclaredFields();
				for (int i = 0; i < teaField.length; i++) {
					String name = teaField[i].getName();
					String value = teaField[i].get(tt).toString();
					int index ;
					if(progressHandle.teaFieldMap.containsKey(name)){
						index = progressHandle.teaFieldMap.get(name); 
					}else{
						index = i;
					}
					rowData[index] = value; 
				}
				writer.addRecord(rowData);
			}
			
			for (int i = 0; i < 9; i++) {
				lineCnt+= size/10;
				publish(new Teacher());
				try {
					Thread.sleep(200);
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
			}
			// 定义输出流，并关联的一个文件
			fos = new FileOutputStream(filename);
			// 写入数据
			writer.write(fos);
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fos.close();
			} catch (Exception e) {
			}
		}
		
		return null;
	}

	@Override
	protected void process(List<Teacher> chunks) {
		// TODO Auto-generated method stub
		if (progressHandle != null) {
			if(size != 0)
				progressHandle.teacherBar.setValue(lineCnt * Values.proMaxSize/ size);
        }
	}

	@Override
	protected void done() {
		// TODO Auto-generated method stub
		if (progressHandle != null) {
            progressHandle.teacherBar.setValue(Values.proMaxSize);
        }
		JOptionPane.showMessageDialog(null, "导出监考数据完毕！",
				"通知", JOptionPane.INFORMATION_MESSAGE);
		progressHandle.teacherBar.setValue(Values.proMinSize);
	}
	
}
