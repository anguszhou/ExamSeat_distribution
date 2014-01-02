package com;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;
import javax.swing.SwingWorker;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class TaskImgdoc extends SwingWorker<List<Student>, Student>{

	private int lineCnt = 0;
	private int preLineCnt = 0;
	private int size = 0;
	private DefaultProgressHandle progressHandle = null;
	final private int columnSize = 10;
	private int docIndex;
	final private int docSize = 50;
	private String outputPath;
	private String prefix;
	private String suffix;
	private String marks;
	private int year;
	
	public TaskImgdoc(String prefix,String suffix,String outputPath, int docIndex,String marks,int year){
		this.outputPath = outputPath;
		this.prefix = prefix;
		this.suffix = suffix;
		this.docIndex = docIndex;
		this.marks = marks;
		this.year = year;
	}

	public void setProgressHandle(DefaultProgressHandle progressHandle) {
		this.progressHandle = progressHandle;
	}
	
	@Override
	protected List<Student> doInBackground() throws Exception {
		// TODO Auto-generated method stub
        
        int pSize = progressHandle.pList.size();
        int index = 0;
        System.out.println("place size: "+pSize);
        if(size == 0){
			size = ExamSeat.getNotZeroClass(progressHandle.pList)*2;
		}
        
        for (int i = 0; i < docIndex * docSize; i++) {
			index += progressHandle.pList.get(i).RS;
		}
        System.out.println("线程："+docIndex+":前置考生数："+index);
        int docNum = (int)Math.ceil((double)pSize / docSize);
        //for (int m = 0; m < docNum; m++) {
    	int m = this.docIndex;
    	DOCWriter writer = new DOCWriter(); 
        writer.createNewDocument();  
        int hasBord = 0;
        
        StringBuffer filename = new StringBuffer();
        filename.append(this.outputPath).append(File.separator).append(year).append("年考生图像对照表(").append(docIndex+1).append(").doc");
        writer.saveAs(filename.toString());
        
        writer.setPageSetup(1, 20, 25, 30, 20);
        writer.setAlignment(1);
        //writer.setVisible(true);
        
        int stuInfo = 4;
        
        for(int i=m*docSize ; i< (m+1)*docSize && i<pSize && progressHandle.pList.get(i).RS > 0; i++){
        	Place pp = progressHandle.pList.get(i);
        	int temp = index;
        	//试场 pp ： 第一单元考试对照表
            StringBuffer sb = new StringBuffer();
            sb.append(pp.SC).append("  地点:").append(pp.JS).append("    第一单元");
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 20);
            writer.insertToDocument(sb.toString());
            
            int cols = columnSize;
            int rows = (int)Math.ceil((double)pp.RS/columnSize);
            
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 8);
            writer.createNewTable(rows*stuInfo, cols, hasBord);
            
        	//循环的行数： table总行数/stuInfo , 每个学生包含5列信息（照片，考生编号，姓名，考试代码，考试科目）
            for(int j=0 ; j<rows ; j++){
            	
            	String[] imgs = new String[columnSize];
            	String[] id = new String[columnSize];
            	String[] name = new String[columnSize];
            	String[] code = new String[columnSize];
            	
            	//循环的列数，每类打印10个考生信息
            	for (int k = 0; k < cols && index < temp+pp.RS; k++) {
            		Student ss = progressHandle.stuList.get(index);
            		
            		try{
            			Integer.valueOf(ss.ZZLLM);
            			
            			StringBuffer imgPath = new StringBuffer();
                		imgPath.append(this.prefix).append(ss.BMH).append(this.suffix);
                		File img = new File(imgPath.toString());
                		if(img.exists()){
                			imgs[k] = imgPath.toString();
                		}
                		id[k] = ss.KSBH;
                		name[k] = ss.XM;
                		code[k] = ss.ZZLLM+ss.ZZLLMC;
            		}catch(NumberFormatException e){
            			k--;
            			//System.out.println("第一单元：教室:"+(i+1)+":"+(index+1)+"，考生："+ss.XM+",没有考试");
            		}
            		index++;
				}
            	writer.insertToTable(imgs, j*stuInfo);
            	writer.insertRowToTable(id, j*stuInfo+1);
            	writer.insertRowToTable(name, j*stuInfo+2);
            	writer.insertRowToTable(code, j*stuInfo+3);
            }
            writer.moveDown(1);
            writer.enterDown(1);
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 12);
            writer.insertToDocument(marks);
            writer.nextPage();
            
          //试场 pp ： 第二单元考试对照表
            index = temp;
            StringBuffer sb2 = new StringBuffer();
            sb2.append(pp.SC).append("  地点:").append(pp.JS).append("    第二单元");
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 20);
            writer.insertToDocument(sb2.toString());
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 8);
            writer.createNewTable(rows*stuInfo, cols, hasBord);
            
        	//循环的行数： table总行数/5 , 每个学生包含5列信息（照片，考生编号，姓名，考试代码，考试科目）
            for(int j=0 ; j<rows ; j++){
            	
            	String[] imgs = new String[columnSize];
            	String[] id = new String[columnSize];
            	String[] name = new String[columnSize];
            	String[] code = new String[columnSize];
            	
            	//循环的列数，每类打印10个考生信息
            	for (int k = 0; k < cols && index < temp+pp.RS; k++) {
            		Student ss = progressHandle.stuList.get(index);
            		
            		try{
            			Integer.valueOf(ss.WGYM);
            			
            			StringBuffer imgPath = new StringBuffer();
                		imgPath.append(this.prefix).append(ss.BMH).append(this.suffix);
                		File img = new File(imgPath.toString());
                		if(img.exists()){
                			imgs[k] = imgPath.toString();
                		}
                		id[k] = ss.KSBH;
                		name[k] = ss.XM;
                		code[k] = ss.WGYM+ss.WGYMC;
            		}catch(NumberFormatException e){
            			k--;
            			//System.out.println("第二单元：教室:"+(i+1)+":"+(index+1)+"，考生："+ss.XM+",没有考试");
            		}
            		index++;
				}
            	writer.insertToTable(imgs, j*stuInfo);
            	writer.insertRowToTable(id, j*stuInfo+1);
            	writer.insertRowToTable(name, j*stuInfo+2);
            	writer.insertRowToTable(code, j*stuInfo+3);
            }
            writer.moveDown(1);
            writer.enterDown(1);
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 12);
            writer.insertToDocument(marks);
            writer.nextPage();
            
          //试场 pp ： 第三单元考试对照表
            index = temp;
            StringBuffer sb3 = new StringBuffer();
            sb3.append(pp.SC).append("  地点:").append(pp.JS).append("    第三单元");
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 20);
            writer.insertToDocument(sb3.toString());
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 8);
            writer.createNewTable(rows*stuInfo, cols, hasBord);
            
        	//循环的行数： table总行数/5 , 每个学生包含5列信息（照片，考生编号，姓名，考试代码，考试科目）
            for(int j=0 ; j<rows ; j++){
            	
            	String[] imgs = new String[columnSize];
            	String[] id = new String[columnSize];
            	String[] name = new String[columnSize];
            	String[] code = new String[columnSize];
            	
            	//循环的列数，每类打印10个考生信息
            	for (int k = 0; k < cols && index < temp+pp.RS; k++) {
            		Student ss = progressHandle.stuList.get(index);
            		
            		try{
            			Integer.valueOf(ss.YWK1M);
            			
            			StringBuffer imgPath = new StringBuffer();
                		imgPath.append(this.prefix).append(ss.BMH).append(this.suffix);
                		File img = new File(imgPath.toString());
                		if(img.exists()){
                			imgs[k] = imgPath.toString();
                		}
                		id[k] = ss.KSBH;
                		name[k] = ss.XM;
                		code[k] = ss.YWK1M+ss.YWK1MC;
            		}catch(NumberFormatException e){
            			k--;
            			//System.out.println("第三单元：教室:"+(i+1)+":"+(index+1)+"，考生："+ss.XM+",没有考试");
            		}
            		index++;
				}
            	writer.insertToTable(imgs, j*stuInfo);
            	writer.insertRowToTable(id, j*stuInfo+1);
            	writer.insertRowToTable(name, j*stuInfo+2);
            	writer.insertRowToTable(code, j*stuInfo+3);
            }
            writer.moveDown(1);
            writer.enterDown(1);
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 12);
            writer.insertToDocument(marks);
            writer.nextPage();
            
          //试场 pp ： 第四单元考试对照表
            index = temp;
            StringBuffer sb4 = new StringBuffer();
            sb4.append(pp.SC).append("  地点:").append(pp.JS).append("    第四单元");
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 20);
            writer.insertToDocument(sb4.toString());
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 8);
            writer.createNewTable(rows*stuInfo, cols, hasBord);
            
        	//循环的行数： table总行数/5 , 每个学生包含5列信息（照片，考生编号，姓名，考试代码，考试科目）
            for(int j=0 ; j<rows ; j++){
            	
            	String[] imgs = new String[columnSize];
            	String[] id = new String[columnSize];
            	String[] name = new String[columnSize];
            	String[] code = new String[columnSize];
            	
            	//循环的列数，每类打印10个考生信息
            	for (int k = 0; k < cols && index < temp+pp.RS; k++) {
            		Student ss = progressHandle.stuList.get(index);
            		
            		try{
            			Integer.valueOf(ss.YWK2M);
            			
            			StringBuffer imgPath = new StringBuffer();
                		imgPath.append(this.prefix).append(ss.BMH).append(this.suffix);
                		File img = new File(imgPath.toString());
                		if(img.exists()){
                			imgs[k] = imgPath.toString();
                		}
                		id[k] = ss.KSBH;
                		name[k] = ss.XM;
                		code[k] = ss.YWK2M+ss.YWK2MC;
            		}catch(NumberFormatException e){
            			k--;
            			//System.out.println("第四单元：教室:"+(i+1)+":"+(index+1)+"，考生："+ss.XM+",没有考试");
            		}
            		index++;
				}
            	writer.insertToTable(imgs, j*stuInfo);
            	writer.insertRowToTable(id, j*stuInfo+1);
            	writer.insertRowToTable(name, j*stuInfo+2);
            	writer.insertRowToTable(code, j*stuInfo+3);
            }
            writer.moveDown(1);
            writer.enterDown(1);
            writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 12);
            writer.insertToDocument(marks);
            if( i+1 < (m+1)*docSize && i+1 < pSize && progressHandle.pList.get(i+1).RS > 0){
            	
            	writer.nextPage();
            }
            lineCnt++;
            publish(new Student());
            if(lineCnt % stuInfo == 0){
            	System.gc();
            }
            writer.save();
            //System.out.println(docIndex+":"+lineCnt+":"+lineCnt * Values.proMaxSize/ size);
        }
        
       /* 
        StringBuffer filename = new StringBuffer();
        filename.append(this.outputPath).append(File.separator).append("ImgDoc(").append(docIndex+1).append(").doc");
        writer.saveAs(filename.toString());
        */
        for (int i = 0; i < writer.getTablesCount(); i++) {
			Dispatch table = writer.getTable(i+1);
			Dispatch rows = Dispatch.call(table, "Rows").toDispatch();
			int rowNum =  Dispatch.get(rows,"Count").getInt();
			
			for (int j = 0; j < rowNum/stuInfo; j++) {
				for (int k = 0; k < columnSize; k++) {
					Dispatch cell = Dispatch.call(table, "Cell", new Variant(j*stuInfo+3),new Variant(k+1)).toDispatch();
					Dispatch.call(cell, "Select");   
					Dispatch.put(writer.getAlignment(), "Alignment", 0);
				}
			}
			
			/*for (int j = 0; j < rowNum/stuInfo; j++) {
				for (int k = 0; k < columnSize; k++) {
					writer.mergeCell2(table,j+1,k+1,j+stuInfo,k+1);
				}
				writer.moveDown(1);
			}*/
			if(i % 4 ==0 ){
				publish(new Student());
			}
			
			/*writer.mergeCell2(table,1,1,8,1);
        	writer.mergeCell2(table,1,2,8,2);
        	writer.mergeCell2(table,2,1,9,1);
        	writer.mergeCell2(table,2,2,9,2);*/
		}
        
        writer.save();
        writer.close();
        writer.quit();
		//}
		return null;
	}

	@Override
	protected void process(List<Student> chunks) {
		// TODO Auto-generated method stub
		double pro = 0;
		//System.out.println("线程："+docIndex+", preLineCnt:"+preLineCnt+",lineCnt:"+lineCnt);
		if (progressHandle != null) {
			/*
			 if(size != 0)
				pro = progressHandle.progressBar.getValue();
				System.out.println("线程："+docIndex+": Preprovalue："+pro+", size:"+size);
				pro = ( pro * size/Values.proMaxSize + (lineCnt-preLineCnt)) *Values.proMaxSize /size;
				if(Math.ceil(pro) < Values.proMaxSize ){
					progressHandle.progressBar.setValue((int)Math.round(pro));
				}else{
					progressHandle.progressBar.setValue((int)Math.ceil(Values.proMaxSize));
				}
				
				System.out.println("线程："+docIndex+": provalue："+progressHandle.progressBar.getValue());
			 */
				progressHandle.addProgress();
        } 
		preLineCnt = lineCnt;
	}

	@Override
	protected void done() {
		// TODO Auto-generated method stub
		/*
		double pro = 0;
		if (progressHandle != null) {
			pro = progressHandle.progressBar.getValue();
			pro = ( pro * size/Values.proMaxSize + (lineCnt-preLineCnt)) *Values.proMaxSize /size;
			if(Math.ceil(pro) < Values.proMaxSize ){
				progressHandle.progressBar.setValue((int)Math.round(pro));
			}else{
				progressHandle.progressBar.setValue((int)Math.ceil(Values.proMaxSize));
			}
        }
        */
		JOptionPane.showMessageDialog(null, "成功生成图像对照word文档！："+(docIndex+1),
				"通知", JOptionPane.INFORMATION_MESSAGE);
		if(progressHandle.progressBar.getValue() == Values.proMaxSize){
			progressHandle.progressBar.setValue(Values.proMinSize);
		}
		//progressHandle.progressBar.setValue(0);
	}
	
	/*
	@Override
	protected List<Student> doInBackground() throws Exception {
		// TODO Auto-generated method stub
        
        int pSize = progressHandle.pList.size();
        int index = 0;
        System.out.println("place size: "+pSize);
        if(size == 0){
			size = pSize;
		}
        
        int docNum = (int)Math.ceil((double)pSize / docSize);
        for (int m = 0; m < docNum; m++) {
        	DOCWriter writer = new DOCWriter(); 
            writer.createNewDocument();    
            writer.setPageSetup(1, 5, 5, 5, 5);
            writer.setAlignment(1);
            //writer.setVisible(true);
            
            for(int i=m*docSize ; i< (m+1)*docSize && i<pSize && progressHandle.pList.get(i).RS > 0; i++){
            	Place pp = progressHandle.pList.get(i);
            	int temp = index;
            	//试场 pp ： 第一单元考试对照表
            	String str = "照片对照表";
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 20);  
                writer.insertToDocument(str);
            	
                StringBuffer sb = new StringBuffer();
                sb.append(pp.SC).append("  :  ").append(pp.JS).append("    第一单元");
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 10);
                writer.insertToDocument(sb.toString());
                
                int cols = columnSize;
                int rows = (int)Math.ceil((double)pp.RS/columnSize);
                
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 8);
                writer.createNewTable(rows*5, cols, 0);
                
            	//循环的行数： table总行数/5 , 每个学生包含5列信息（照片，考生编号，姓名，考试代码，考试科目）
                for(int j=0 ; j<rows ; j++){
                	
                	String[] imgs = new String[columnSize];
                	String[] id = new String[columnSize];
                	String[] name = new String[columnSize];
                	String[] code = new String[columnSize];
                	String[] exam = new String[columnSize];
                	
                	//循环的列数，每类打印10个考生信息
                	for (int k = 0; k < cols && index < temp+pp.RS; k++) {
                		Student ss = progressHandle.stuList.get(index);
                		
                		try{
                			Integer.valueOf(ss.ZZLLM);
                			
                			StringBuffer imgPath = new StringBuffer();
                    		imgPath.append(this.prefix).append(ss.BMH).append(this.suffix);
                    		File img = new File(imgPath.toString());
                    		if(img.exists()){
                    			imgs[k] = imgPath.toString();
                    		}
                    		id[k] = ss.KSBH;
                    		name[k] = ss.XM;
                    		code[k] = ss.ZZLLM;
                    		exam[k] = ss.ZZLLMC;
                		}catch(NumberFormatException e){
                			System.out.println("第一单元："+index+"，考生："+ss.XM+",没有考试");
                		}
                		index++;
					}
                	writer.insertToTable(imgs, j*5);
                	writer.insertRowToTable(id, j*5+1);
                	writer.insertRowToTable(name, j*5+2);
                	writer.insertRowToTable(code, j*5+3);
                	writer.insertRowToTable(exam, j*5+4);
                }
                writer.moveDown(1);
                writer.enterDown(1);
                writer.nextPage();
                
              //试场 pp ： 第二单元考试对照表
                index = temp;
                String str2 = "照片对照表";
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 20);  
                writer.insertToDocument(str2);
            	
                StringBuffer sb2 = new StringBuffer();
                sb2.append(pp.SC).append("  :  ").append(pp.JS).append("    第二单元");
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 10);
                writer.insertToDocument(sb2.toString());
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 8);
                writer.createNewTable(rows*5, cols, 0);
                
            	//循环的行数： table总行数/5 , 每个学生包含5列信息（照片，考生编号，姓名，考试代码，考试科目）
                for(int j=0 ; j<rows ; j++){
                	
                	String[] imgs = new String[columnSize];
                	String[] id = new String[columnSize];
                	String[] name = new String[columnSize];
                	String[] code = new String[columnSize];
                	String[] exam = new String[columnSize];
                	
                	//循环的列数，每类打印10个考生信息
                	for (int k = 0; k < cols && index < temp+pp.RS; k++) {
                		Student ss = progressHandle.stuList.get(index);
                		
                		try{
                			Integer.valueOf(ss.WGYM);
                			
                			StringBuffer imgPath = new StringBuffer();
                    		imgPath.append(this.prefix).append(ss.BMH).append(this.suffix);
                    		File img = new File(imgPath.toString());
                    		if(img.exists()){
                    			imgs[k] = imgPath.toString();
                    		}
                    		id[k] = ss.KSBH;
                    		name[k] = ss.XM;
                    		code[k] = ss.WGYM;
                    		exam[k] = ss.WGYMC;
                		}catch(NumberFormatException e){
                			System.out.println("第二单元："+index+"，考生："+ss.XM+",没有考试");
                		}
                		index++;
					}
                	writer.insertToTable(imgs, j*5);
                	writer.insertRowToTable(id, j*5+1);
                	writer.insertRowToTable(name, j*5+2);
                	writer.insertRowToTable(code, j*5+3);
                	writer.insertRowToTable(exam, j*5+4);
                }
                writer.moveDown(1);
                writer.enterDown(1);
                writer.nextPage();
                
              //试场 pp ： 第三单元考试对照表
                index = temp;
                String str3 = "照片对照表";
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 20);  
                writer.insertToDocument(str3);
            	
                StringBuffer sb3 = new StringBuffer();
                sb3.append(pp.SC).append("  :  ").append(pp.JS).append("    第三单元");
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 10);
                writer.insertToDocument(sb3.toString());
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 8);
                writer.createNewTable(rows*5, cols, 0);
                
            	//循环的行数： table总行数/5 , 每个学生包含5列信息（照片，考生编号，姓名，考试代码，考试科目）
                for(int j=0 ; j<rows ; j++){
                	
                	String[] imgs = new String[columnSize];
                	String[] id = new String[columnSize];
                	String[] name = new String[columnSize];
                	String[] code = new String[columnSize];
                	String[] exam = new String[columnSize];
                	
                	//循环的列数，每类打印10个考生信息
                	for (int k = 0; k < cols && index < temp+pp.RS; k++) {
                		Student ss = progressHandle.stuList.get(index);
                		
                		try{
                			Integer.valueOf(ss.YWK1M);
                			
                			StringBuffer imgPath = new StringBuffer();
                    		imgPath.append(this.prefix).append(ss.BMH).append(this.suffix);
                    		File img = new File(imgPath.toString());
                    		if(img.exists()){
                    			imgs[k] = imgPath.toString();
                    		}
                    		id[k] = ss.KSBH;
                    		name[k] = ss.XM;
                    		code[k] = ss.YWK1M;
                    		exam[k] = ss.YWK1MC;
                		}catch(NumberFormatException e){
                			System.out.println("第三单元："+index+"，考生："+ss.XM+",没有考试");
                		}
                		index++;
					}
                	writer.insertToTable(imgs, j*5);
                	writer.insertRowToTable(id, j*5+1);
                	writer.insertRowToTable(name, j*5+2);
                	writer.insertRowToTable(code, j*5+3);
                	writer.insertRowToTable(exam, j*5+4);
                }
                writer.moveDown(1);
                writer.enterDown(1);
                writer.nextPage();
                
              //试场 pp ： 第四单元考试对照表
                index = temp;
                String str4 = "照片对照表";
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 20);  
                writer.insertToDocument(str4);
            	
                StringBuffer sb4 = new StringBuffer();
                sb4.append(pp.SC).append("  :  ").append(pp.JS).append("    第四单元");
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 10);
                writer.insertToDocument(sb4.toString());
                writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 8);
                writer.createNewTable(rows*5, cols, 0);
                
            	//循环的行数： table总行数/5 , 每个学生包含5列信息（照片，考生编号，姓名，考试代码，考试科目）
                for(int j=0 ; j<rows ; j++){
                	
                	String[] imgs = new String[columnSize];
                	String[] id = new String[columnSize];
                	String[] name = new String[columnSize];
                	String[] code = new String[columnSize];
                	String[] exam = new String[columnSize];
                	
                	//循环的列数，每类打印10个考生信息
                	for (int k = 0; k < cols && index < temp+pp.RS; k++) {
                		Student ss = progressHandle.stuList.get(index);
                		
                		try{
                			Integer.valueOf(ss.YWK2M);
                			
                			StringBuffer imgPath = new StringBuffer();
                    		imgPath.append(this.prefix).append(ss.BMH).append(this.suffix);
                    		File img = new File(imgPath.toString());
                    		if(img.exists()){
                    			imgs[k] = imgPath.toString();
                    		}
                    		id[k] = ss.KSBH;
                    		name[k] = ss.XM;
                    		code[k] = ss.YWK2M;
                    		exam[k] = ss.YWK2MC;
                		}catch(NumberFormatException e){
                			System.out.println("第四单元："+index+"，考生："+ss.XM+",没有考试");
                		}
                		index++;
					}
                	writer.insertToTable(imgs, j*5);
                	writer.insertRowToTable(id, j*5+1);
                	writer.insertRowToTable(name, j*5+2);
                	writer.insertRowToTable(code, j*5+3);
                	writer.insertRowToTable(exam, j*5+4);
                }
                if( i+1 < (m+1)*docSize && i+1 < pSize && progressHandle.pList.get(i+1).RS > 0){
                	writer.moveDown(1);
                    writer.enterDown(1);
                	writer.nextPage();
                }
                lineCnt++;
                publish(new Student());
                if(lineCnt % 5 == 0){
                	System.gc();
                }
                System.out.println(lineCnt+":"+lineCnt * 100/ size);
            }
            
            StringBuffer filename = new StringBuffer();
            filename.append(this.outputPath).append(File.separator).append("ImgDoc(").append(m+1).append(").doc");
            writer.saveAs(filename.toString());
            writer.close();
            writer.quit();
		}
		return null;
	}
	*/
	
}
