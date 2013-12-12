package com;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JProgressBar;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

public class DefaultProgressHandle{

	JProgressBar progressBar = null;
	int lineCnt=0, size=0;
	JTable table = null;
	DefaultTableModel model;
	int count = 0;
	List<Student> stuList = null;
	List<Place> pList = null;
	List<String> field= new ArrayList<String>();
	Map<String,Integer> fieldMap = new HashMap<String,Integer>();
	
	public void setProgressBar(JProgressBar progressBar) {
	        this.progressBar = progressBar;	        
    }
	public void setTable(JTable table) {
		this.table = table;
		model = (DefaultTableModel) table.getModel();
	}
	
	public void processInProgress(List<Student>stuList,int progress){
		// TODO Auto-generated method stub
		for(Student ss : stuList){
			Object[] newRow = {++count,ss.KSBH, ss.XM,ss.SC,ss.JS,ss.ZZLLMC,
					ss.WGYMC,ss.YWK1MC,ss.YWK2MC};
			model.addRow(newRow);
		}

		progressBar.setValue(progress);
	}
	
	synchronized public void addProgress(){
		lineCnt++;
		System.out.println("progress, lineCnt:"+lineCnt+",Size:"+size+",pro:"+lineCnt * Values.proMaxSize / size);
		this.progressBar.setValue(lineCnt * Values.proMaxSize / size);
	}
	
	public void processComplete(List<Student>stuList) {
		// TODO Auto-generated method stub
		progressBar.setValue(progressBar.getMaximum());
	}

}
