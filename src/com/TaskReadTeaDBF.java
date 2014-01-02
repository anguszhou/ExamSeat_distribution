package com;

import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutionException;

import javax.swing.JOptionPane;
import javax.swing.SwingWorker;

import com.linuxense.javadbf.DBFField;
import com.linuxense.javadbf.DBFReader;

public class TaskReadTeaDBF extends SwingWorker<List<Teacher>, Teacher> {

	String filepath;
	private int lineCnt = 0;
	private int size = 0;
	private DefaultProgressHandle progressHandle = null;
	
	public TaskReadTeaDBF(String filepath){
		this.filepath = filepath;
	}

	public void setProgressHandle(DefaultProgressHandle progressHandle) {
		this.progressHandle = progressHandle;
	}

	
	@Override
	protected List<Teacher> doInBackground() throws Exception {
		// TODO Auto-generated method stub
		InputStream fis = null;
		List<Teacher> teaList = new ArrayList<Teacher>();
		try {
			// 读取文件的输入流
			fis = new FileInputStream(filepath);
			// 根据输入流初始化一个DBFReader实例，用来读取DBF文件信息
			DBFReader reader = new DBFReader(fis);
			reader.setCharactersetName("GBK");
			// 调用DBFReader对实例方法得到path文件中字段的个数
			int fieldsCount = reader.getFieldCount();
			// 取出字段信息
			for (int i = 0; i < fieldsCount; i++) {
				DBFField field = reader.getField(i);
				progressHandle.teaField.add(field.getName());
				progressHandle.teaFieldMap.put(field.getName(), i);
			}
			Object[] rowValues;
			// 一条条取出path文件中记录
			if(size == 0){
				size = reader.getRecordCount();
			}
			while ((rowValues = reader.nextRecord()) != null) {
				Teacher tea = new Teacher();
				tea = setTeacher(lineCnt,rowValues);
				//stu.print();
				lineCnt++;
				publish(tea);
				teaList.add(tea);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fis.close();
			}catch (Exception e) {}
		}
		
		return teaList;
	}
	@Override
	protected void process(List<Teacher> chunks) {
		// TODO Auto-generated method stub
		if (progressHandle != null) {
			if(size != 0){
				progressHandle.processInTea(chunks,lineCnt * Values.proMaxSize/ size);
			}
					
        }
	}
	@Override
	protected void done() {
		try {
			if (progressHandle != null) {
		            progressHandle.teacherBar.setValue(Values.proMaxSize);
		        }
			JOptionPane.showMessageDialog(null, "导入数据成功！",
					"通知", JOptionPane.INFORMATION_MESSAGE);
			progressHandle.teacherBar.setValue(Values.proMinSize);
			progressHandle.teaList = get();
			
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ExecutionException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	private Teacher setTeacher(int id , Object[] rowValues) throws Exception, Exception{
		Teacher tea = new Teacher();
		Field[] teaField = tea.getClass().getDeclaredFields();
		for (int i = 0; i < teaField.length; i++) {
			String name = teaField[i].getName();
			if(progressHandle.teaFieldMap.containsKey(name)){
				teaField[i].set(tea, rowValues[progressHandle.teaFieldMap.get(name)].toString().trim());
			}else{
				tea.id = id;
			}
		}
		return tea;
	}
}
