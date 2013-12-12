package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.linuxense.javadbf.DBFField;
import com.linuxense.javadbf.DBFReader;
import com.linuxense.javadbf.DBFWriter;

public class DBFHandler {
	
	public static List<Student> readDBF(String path) {
		InputStream fis = null;
		List<Student> stuList = new ArrayList<Student>();
		List<String> fields = new ArrayList<String>();
		Map<String , List> data = new HashMap<String , List>();
		try {
			// 读取文件的输入流
			fis = new FileInputStream(path);
			// 根据输入流初始化一个DBFReader实例，用来读取DBF文件信息
			DBFReader reader = new DBFReader(fis);
			reader.setCharactersetName("GBK");
			// 调用DBFReader对实例方法得到path文件中字段的个数
			int fieldsCount = reader.getFieldCount();
			// 取出字段信息
			for (int i = 0; i < fieldsCount; i++) {
				DBFField field = reader.getField(i);
				fields.add(field.getName());
				//System.out.println(field.getName());
			}
			//System.exit(0);
			Object[] rowValues;
			// 一条条取出path文件中记录
			while ((rowValues = reader.nextRecord()) != null) {
				Student stu = new Student();
				/*stu.setId(rowValues[4].toString().trim());
				stu.setName(rowValues[2].toString().trim());
				stu.setTest1(rowValues[53].toString().trim());
				stu.setTest2(rowValues[55].toString().trim());
				stu.setTest3(rowValues[57].toString().trim());
				stu.setTest4(rowValues[59].toString().trim());*/
				//stu.print();
				//stu.setImg();
				stuList.add(stu);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fis.close();
			} catch (Exception e) {
			}
		}
		System.out.println(stuList.size());
		data.put("fields", fields);
		data.put("stu", stuList);
		return stuList;
	}
	
	public static void writeDBF(String path) {
		OutputStream fos = null;
		try {
			// 定义DBF文件字段
			DBFField[] fields = new DBFField[3];
			// 分别定义各个字段信息，setFieldName和setName作用相同,
			// 只是setFieldName已经不建议使用
			fields[0] = new DBFField();
			// fields[0].setFieldName("emp_code");
			fields[0].setName("semp_code");
			fields[0].setDataType(DBFField.FIELD_TYPE_C);
			fields[0].setFieldLength(10);
			fields[1] = new DBFField();
			// fields[1].setFieldName("emp_name");
			fields[1].setName("emp_name");
			fields[1].setDataType(DBFField.FIELD_TYPE_C);
			fields[1].setFieldLength(20);
			fields[2] = new DBFField();
			// fields[2].setFieldName("salary");
			fields[2].setName("salary");
			fields[2].setDataType(DBFField.FIELD_TYPE_N);
			fields[2].setFieldLength(12);
			fields[2].setDecimalCount(2);
			DBFWriter writer = new DBFWriter(new File(path));
			// 定义DBFWriter实例用来写DBF文件
			//DBFWriter writer = new DBFWriter();
			// 把字段信息写入DBFWriter实例，即定义表结构
			
			writer.setFields(fields);
			// 一条条的写入记录
			Object[] rowData = new Object[3];
			rowData[0] = "1000";
			rowData[1] = "John";
			rowData[2] = new Double(5000.00);
			writer.addRecord(rowData);
			rowData = new Object[3];
			rowData[0] = "1001";
			rowData[1] = "Lalit";
			rowData[2] = new Double(3400.00);
			writer.addRecord(rowData);
			rowData = new Object[3];
			rowData[0] = "1002";
			rowData[1] = "Rohit";
			rowData[2] = new Double(7350.00);
			writer.addRecord(rowData);
			// 定义输出流，并关联的一个文件
			fos = new FileOutputStream(path);
			// 写入数据
			//writer.write(fos);
			writer.write();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fos.close();
			} catch (Exception e) {
			}
		}
	}
	
	
	
	public static void main(String[] args) {
		writeDBF("E:\\google下载\\2014准确数据aaaaaaaa\\12.dbf");
	}
}
