package com;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;
import javax.swing.SwingWorker;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class TaskTeadoc extends SwingWorker<List<Teacher>, Teacher>{

	private int lineCnt = 0;
	private int size = 0;
	private DefaultProgressHandle progressHandle = null;
	private int docIndex;
	private String outputPath;
	private String imgPath;
	private int year;
	private String type;
	
	public TaskTeadoc(String imgPath,String outputPath, int year,String type){
		this.outputPath = outputPath;
		this.imgPath = imgPath;
		this.year = year;
		this.type = type;
	}

	public void setProgressHandle(DefaultProgressHandle progressHandle) {
		this.progressHandle = progressHandle;
	}
	
	@Override
	protected List<Teacher> doInBackground() throws Exception {
		// TODO Auto-generated method stub
		 StringBuffer dirPath = new StringBuffer();
		 dirPath.append(imgPath).append(File.separator).append("tmpImg");
		 File dir = new File(dirPath.toString());
		 if(!dir.isDirectory()){
			 dir.mkdir();
		 }
		
		
		StringBuffer sb = new StringBuffer();
		sb.append(outputPath).append(File.separator).append(year).append(type).append("按排文档.doc");
		
		DOCWriter writer = new DOCWriter(); 
        writer.createNewDocument(); 
        writer.saveAs(sb.toString()); 
        writer.setPageSetup(0, 50, 60, 20, 20);
        writer.setAlignment(1);
		
        sb = new StringBuffer();
        sb.append("西安交大").append(year).append("年硕士生入学考试");
        
        writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 15);
        int rowSize = 16;
        int colSize = 2;
        int perSize = 4;
        int num = 8;
        int page = (int)Math.ceil((double)progressHandle.teaList.size()/perSize);
        int index = 0;
        if(size == 0 ){
        	size = page * 2 + 2;
        }
        for (int i = 0; i < page ; i++) {
        	writer.createNewTable(rowSize, colSize, 1);
        	
        	for (int j = 0; j < rowSize/num; j++) {
        		
            	String[] id = new String[colSize];
            	String[] head =new String[colSize];
            	String[] zw = new String[colSize];
            	String[] imgs = new String[colSize];
            	String[] dw = new String[colSize];
            	String[] xm = new String[colSize];
            	String[] sc = new String[colSize];
            	String[] dd = new String[colSize];
        		
				for (int k = 0; k < colSize && index < progressHandle.teaList.size(); k++) {
					 Teacher tmp = progressHandle.teaList.get(index);
					 index++;
					 id[k] = "编号:"+tmp.KH;
					 head[k] = sb.toString();
					 zw[k] = " "+tmp.ZW;
					 
					 String zpValue = tmp.ZP;
					 if(zpValue== null || zpValue.length()<1){
						 zpValue = String.valueOf(tmp.id);
					 }
					 StringBuffer img = new StringBuffer();
					 img.append(imgPath).append(File.separator).append(zpValue);
					 File zp = new File(img.toString());
					 if(!zp.exists()){
						 StringBuffer ss = new StringBuffer();
						 ss.append(dir.getAbsolutePath()).append(File.separator).append(zpValue);
						 createJPG.create(90, 120, ss.toString());
						 imgs[k] = ss.toString();
					 }else{
						 imgs[k] = img.toString();
					 }
					 System.out.println(imgs[k]);
					 dw[k] = "工作单位:"+tmp.DW;
					 xm[k] = "姓名:"+tmp.XM;
					 sc[k] = "试场:"+tmp.SC;
					 dd[k] = "地点:"+tmp.DD;
				}
				
            	writer.insertRowToTable(id, j*num);
            	writer.insertRowToTable(head, j*num+1);
            	writer.insertRowToTable(zw, j*num+2);
            	writer.insertToTable2(imgs, j*num+3);
            	writer.insertRowToTable(dw, j*num+4);
            	writer.insertRowToTable(xm, j*num+5);
            	writer.insertRowToTable(sc, j*num+6);
            	writer.insertRowToTable(dd, j*num+7);
			}
        	
        	 writer.moveDown(1);
             writer.enterDown(1);
        	 writer.nextPage();  
        	 lineCnt++;
        	 publish(new Teacher());
		}
        writer.save();
        
        for (int i = 0; i < writer.getTablesCount(); i++) {
			Dispatch table = writer.getTable(i+1);
			/*
			Dispatch cell = Dispatch.call(table, "Cell", new Variant(6),new Variant(1)).toDispatch();
			Dispatch.call(cell, "Select");   
			Dispatch.put(writer.getFont(), "Size", 15);  
			Dispatch.put(writer.getAlignment(), "Alignment", 0);
			*/
			writer.mergeCell2(table,1,1,8,1);
        	writer.mergeCell2(table,1,2,8,2);
        	writer.mergeCell2(table,2,1,9,1);
        	writer.mergeCell2(table,2,2,9,2);
        	lineCnt++;
       	 	publish(new Teacher());
		}
        writer.save();
        writer.close();
        writer.quit();
        DeleteFile.deleteDirectory(dir.getAbsolutePath());
		return null;
	}

	@Override
	protected void process(List<Teacher> chunks) {
		// TODO Auto-generated method stub
		//System.out.println("线程："+docIndex+", preLineCnt:"+preLineCnt+",lineCnt:"+lineCnt);
		if (progressHandle != null) {
			if(size != 0)
				progressHandle.teacherBar.setValue(lineCnt * Values.proMaxSize/ size);
        } 
	}

	@Override
	protected void done() {
		
		if (progressHandle != null) {
            progressHandle.teacherBar.setValue(Values.proMaxSize);
        }
		
		JOptionPane.showMessageDialog(null, "成功生成监考按排word文档！",
				"通知", JOptionPane.INFORMATION_MESSAGE);
		
		progressHandle.teacherBar.setValue(Values.proMinSize);
	}
	
}
