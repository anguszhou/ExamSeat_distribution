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

public class TaskDistribute extends SwingWorker<List<Student>, Student>{

	String filepath;
	private int lineCnt = 0;
	private int size = 0;
	private DefaultProgressHandle progressHandle = null;
	
	public TaskDistribute(String filepath){
		this.filepath = filepath;
	}
	public void setProgressHandle(DefaultProgressHandle progressHandle) {
		this.progressHandle = progressHandle;
	}
	@Override
	protected List<Student> doInBackground() throws Exception {
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
			int stuSize = progressHandle.stuList.size();
			int index = 0;
			System.out.println("stuL:"+stuSize);
			
			List<Place> pList = new ArrayList<Place>();
			
			while ((rowValues = reader.nextRecord()) != null) {
				String sc = rowValues[0].toString().trim();
				String js = rowValues[1].toString().trim();
				Double num = Double.valueOf(rowValues[2].toString().trim());
				String bz = rowValues[4].toString().trim();
				
				Place pp = new Place();
				pp.SC = sc;
				pp.JS = js;
				pp.RS = num;
				pp.BZ = bz;
				pList.add(pp);
				
				int temp = index;
				for (int i = index; i < num+temp && index < stuSize; i++) {
					progressHandle.stuList.get(index).SC = sc;
					progressHandle.stuList.get(index).JS = js;
					progressHandle.table.setValueAt(sc, index, 3);
					progressHandle.table.setValueAt(js, index, 4);
					index++;
				}
				
			}
			progressHandle.pList = pList;
			writeDBF(progressHandle.stuList,filepath);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fis.close();
			}catch (Exception e) {}
		}
		return null;
	}

	private void writeDBF(List<Student> stuList,String filepath) {
		// TODO Auto-generated method stub
		String path = filepath.substring(0, filepath.lastIndexOf(File.separator))+File.separator+"2014bxap(已分配考场).dbf";
		System.out.println(path);
		OutputStream fos = null;
		try {
			// 定义DBF文件字段
			// 分别定义各个字段信息，setFieldName和setName作用相同,
			int size = progressHandle.field.size()+2;
			DBFField[] fields = new DBFField[size];
			for (int i = 0; i < size-2; i++) {
				fields[i] = new DBFField();
				fields[i].setName(progressHandle.field.get(i));
				fields[i].setDataType(DBFField.FIELD_TYPE_C);
			}
			
			fields[0].setFieldLength(4);
			fields[1].setFieldLength(9);
			fields[2].setFieldLength(20);
			fields[3].setFieldLength(40);
			fields[4].setFieldLength(15);
			fields[5].setFieldLength(2);
			fields[6].setFieldLength(18);
			fields[7].setFieldLength(8);
			fields[8].setFieldLength(2);
			fields[9].setFieldLength(1);
			fields[10].setFieldLength(1);
			fields[11].setFieldLength(1);
			fields[12].setFieldLength(2);
			fields[13].setFieldLength(6);
			fields[14].setFieldLength(6);
			fields[15].setFieldLength(6);
			fields[16].setFieldLength(60);
			fields[17].setFieldLength(6);
			fields[18].setFieldLength(60);
			fields[19].setFieldLength(80);
			fields[20].setFieldLength(6);
			fields[21].setFieldLength(60);
			fields[22].setFieldLength(220);
			fields[23].setFieldLength(220);
			fields[24].setFieldLength(254);
			fields[25].setFieldLength(220);
			fields[26].setFieldLength(80);
			fields[27].setFieldLength(6);
			fields[28].setFieldLength(40);
			fields[29].setFieldLength(12);
			fields[30].setFieldLength(30);
			fields[31].setFieldLength(1);
			fields[32].setFieldLength(5);
			fields[33].setFieldLength(34);
			fields[34].setFieldLength(6);
			fields[35].setFieldLength(40);
			fields[36].setFieldLength(1);
			fields[37].setFieldLength(1);
			fields[38].setFieldLength(18);
			fields[39].setFieldLength(6);
			fields[40].setFieldLength(18);
			fields[41].setFieldLength(1);
			fields[42].setFieldLength(20);
			fields[43].setFieldLength(5);
			fields[44].setFieldLength(6);
			fields[45].setFieldLength(2);
			fields[46].setFieldLength(1);
			fields[47].setFieldLength(2);
			fields[48].setFieldLength(6);
			fields[49].setFieldLength(60);
			fields[50].setFieldLength(3);
			fields[51].setFieldLength(2);
			fields[52].setFieldLength(3);
			fields[53].setFieldLength(20);
			fields[54].setFieldLength(3);
			fields[55].setFieldLength(20);
			fields[56].setFieldLength(3);
			fields[57].setFieldLength(40);
			fields[58].setFieldLength(3);
			fields[59].setFieldLength(40);
			fields[60].setFieldLength(50);
			fields[61].setFieldLength(50);
			fields[62].setFieldLength(100);
			fields[63].setFieldLength(254);
			fields[64].setFieldLength(100);
			fields[65].setFieldLength(100);
			fields[66].setFieldLength(100);
			fields[67].setFieldLength(100);
			fields[68].setFieldLength(100);
			fields[69].setFieldLength(100);
			fields[70].setFieldLength(100);
			fields[71].setFieldLength(100);
			fields[72].setFieldLength(100);
			
			fields[size-2] = new DBFField();
			fields[size-2].setName("SC");
			fields[size-2].setDataType(DBFField.FIELD_TYPE_C);
			fields[size-2].setFieldLength(20);
			fields[size-1] = new DBFField();
			fields[size-1].setName("JS");
			fields[size-1].setDataType(DBFField.FIELD_TYPE_C);
			fields[size-1].setFieldLength(50);
			
			//DBFWriter writer = new DBFWriter(new File(path));
			// 定义DBFWriter实例用来写DBF文件
			DBFWriter writer = new DBFWriter();
			writer.setCharactersetName("GBK");
			// 把字段信息写入DBFWriter实例，即定义表结构
			writer.setFields(fields);
			// 一条条的写入记录
			Object[] rowData ;
			for(Student ss : stuList){
				rowData = new Object[size];
				
				Field[] stuField = ss.getClass().getDeclaredFields();
				for (int i = 0; i < stuField.length; i++) {
					String name = stuField[i].getName();
					String value = stuField[i].get(ss).toString();
					int index ;
					if(progressHandle.fieldMap.containsKey(name)){
						index = progressHandle.fieldMap.get(name); 
					}else{
						index = i;
					}
					rowData[index] = value; 
				}
				publish(ss);
				writer.addRecord(rowData);
				if(lineCnt*2 <= this.size){
					lineCnt++;
				}
				
			}
			
			// 定义输出流，并关联的一个文件
			fos = new FileOutputStream(path);
			// 写入数据
			writer.write(fos);
			lineCnt = this.size;
			//writer.write();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fos.close();
			} catch (Exception e) {
			}
		}
	}
	@Override
	protected void process(List<Student> chunks) {
		// TODO Auto-generated method stub
		if (progressHandle != null) {
			if(size != 0)
				progressHandle.progressBar.setValue(lineCnt * Values.proMaxSize/ size);
        }
	}

	@Override
	protected void done() {
		// TODO Auto-generated method stub
		if (progressHandle != null) {
            progressHandle.progressBar.setValue(Values.proMaxSize);
        }
		JOptionPane.showMessageDialog(null, "分配考场完毕",
				"通知", JOptionPane.INFORMATION_MESSAGE);
		progressHandle.progressBar.setValue(Values.proMinSize);
	}
	
}
