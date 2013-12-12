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

public class TaskReadDBF extends SwingWorker<List<Student>, Student> {

	String filepath;
	private int lineCnt = 0;
	private int size = 0;
	private DefaultProgressHandle progressHandle = null;
	
	public TaskReadDBF(String filepath){
		this.filepath = filepath;
	}

	public void setProgressHandle(DefaultProgressHandle progressHandle) {
		this.progressHandle = progressHandle;
	}

	
	@Override
	protected List<Student> doInBackground() throws Exception {
		// TODO Auto-generated method stub
		InputStream fis = null;
		List<Student> stuList = new ArrayList<Student>();
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
				progressHandle.field.add(field.getName());
				progressHandle.fieldMap.put(field.getName(), i);
			}
			Object[] rowValues;
			// 一条条取出path文件中记录
			if(size == 0){
				size = reader.getRecordCount();
			}
			while ((rowValues = reader.nextRecord()) != null) {
				Student stu = new Student();
				stu = setStuData(rowValues);
				
				//stu.print();
				lineCnt++;
				publish(stu);
				stuList.add(stu);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fis.close();
			}catch (Exception e) {}
		}
		
		return stuList;
	}
	@Override
	protected void process(List<Student> chunks) {
		// TODO Auto-generated method stub
		if (progressHandle != null) {
			if(size != 0)
				//System.out.println(lineCnt * Values.proMaxSize/ size);
				progressHandle.processInProgress(chunks,lineCnt * Values.proMaxSize/ size);
        }
	}
	@Override
	protected void done() {
		try {
			if (progressHandle != null) {
		            progressHandle.processComplete(get());
		        }
			JOptionPane.showMessageDialog(null, "导入数据成功",
					"通知", JOptionPane.INFORMATION_MESSAGE);
			progressHandle.progressBar.setValue(Values.proMinSize);
			progressHandle.stuList = get();
			
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ExecutionException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	private Student setStuData(Object[] rowValues) throws Exception, Exception{
		Student stu = new Student();
		Field[] stuField = stu.getClass().getDeclaredFields();
		for (int i = 0; i < stuField.length; i++) {
			String name = stuField[i].getName();
			if(progressHandle.fieldMap.containsKey(name)){
				stuField[i].set(stu, rowValues[progressHandle.fieldMap.get(name)].toString().trim());
			}
		}
		return stu;
	}
}
