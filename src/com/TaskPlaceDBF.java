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

public class TaskPlaceDBF extends SwingWorker<List<Place>, Place>{

	String filepath;
	private int lineCnt = 0;
	private int size = 0;
	private DefaultProgressHandle progressHandle = null;
	
	public TaskPlaceDBF(String filepath){
		this.filepath = filepath;
	}
	public void setProgressHandle(DefaultProgressHandle progressHandle) {
		this.progressHandle = progressHandle;
	}
	@Override
	protected List<Place> doInBackground() throws Exception {
		// TODO Auto-generated method stub
		InputStream fis = null;
		try {
			// 读取文件的输入流
			fis = new FileInputStream(filepath);
			// 根据输入流初始化一个DBFReader实例，用来读取DBF文件信息
			DBFReader reader = new DBFReader(fis);
			reader.setCharactersetName("GBK");
			// 调用DBFReader对实例方法得到path文件中字段的个数
			
			Object[] rowValues;
			// 一条条取出path文件中记录
			if(size == 0){
				size = reader.getRecordCount();
			}
			System.out.println("palce:"+size);
			
			List<Place> pList = new ArrayList<Place>();
			
			while ((rowValues = reader.nextRecord()) != null) {
				String sc = rowValues[0].toString().trim();
				String js = rowValues[1].toString().trim();
				Double num = Double.valueOf(rowValues[2].toString().trim());
				String bz = rowValues[4].toString().trim();
				
				Place pp = new Place();
				pp.id = lineCnt;
				pp.SC = sc;
				pp.JS = js;
				pp.RS = num;
				pp.BZ = bz;
				
				if(pp.RS > 0){
					StringBuffer sb = new StringBuffer();
					sb.append("[").append(lineCnt+1).append("] ").append(pp.SC).append(" [人数 : ").append((int)pp.RS).append("] [类型 : ").append(pp.BZ).append("]");
					progressHandle.placeBoxMode.addElement(sb.toString());
					
					pList.add(pp);
					lineCnt++;
					publish(pp);
				}
			}
			progressHandle.pTeaList = pList;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fis.close();
			}catch (Exception e) {}
		}
		return null;
	}

	@Override
	protected void process(List<Place> chunks) {
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
		JOptionPane.showMessageDialog(null, "成功导入考场信息！",
				"通知", JOptionPane.INFORMATION_MESSAGE);
		progressHandle.teacherBar.setValue(Values.proMinSize);
	}
	
}
