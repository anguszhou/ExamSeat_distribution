package com;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JProgressBar;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

public class DefaultProgressHandle{

	JProgressBar progressBar = null , teacherBar;
	int lineCnt=0, size=0;
	JTable table = null , disedPlaceTable , nDisTeaTable;
	DefaultTableModel model , disedPlaceTableModel, nDisTeaTableModel;
	JComboBox placeBox;
	DefaultComboBoxModel placeBoxMode;
	int count = 0 , teaCount = 0;
	List<Student> stuList = null;
	List<Place> pList = null;
	List<Place> pTeaList = null;
	List<Teacher> teaList = null;
	List<String> field= new ArrayList<String>();
	Map<String,Integer> fieldMap = new HashMap<String,Integer>();
	
	List<String> teaField= new ArrayList<String>();
	Map<String,Integer> teaFieldMap = new HashMap<String,Integer>();
	
	public void setProgressBar(JProgressBar progressBar , JProgressBar teacherBar) {
	        this.progressBar = progressBar;
	        this.teacherBar = teacherBar;
    }
	
	public void setCombBox(JComboBox placeBox){
		this.placeBox = placeBox;
		placeBoxMode = (DefaultComboBoxModel) placeBox.getModel();
	}
	public void setTable(JTable table , JTable disedPlaceTable , JTable nDisTeaTable) {
		this.table = table;
		model = (DefaultTableModel) table.getModel();
		
		this.disedPlaceTable = disedPlaceTable;
		disedPlaceTableModel = (DefaultTableModel) disedPlaceTable.getModel();
		
		this.nDisTeaTable = nDisTeaTable;
		nDisTeaTableModel = (DefaultTableModel) nDisTeaTable.getModel();
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
	
	public void processInTea(List<Teacher>teaList,int progress){
		// TODO Auto-generated method stub
		for(Teacher tea : teaList){
			Object[] newRow = {tea.id+1 , tea.KH, tea.XM, tea.XB ,tea.ZW ,tea.DW, tea.SF};
			nDisTeaTableModel.addRow(newRow);
		}

		teacherBar.setValue(progress);
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
