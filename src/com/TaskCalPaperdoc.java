package com;

import java.io.File;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import javax.swing.JOptionPane;
import javax.swing.SwingWorker;

public class TaskCalPaperdoc extends SwingWorker<List<Student>, Student>{

	private int lineCnt = 0;
	private int size = 0;
	private DefaultProgressHandle progressHandle = null;
	private String outputPath;
	private int year;
	
	public TaskCalPaperdoc(String outputPath , int year){
		this.outputPath= outputPath;
		this.year = year;
	}

	public void setProgressHandle(DefaultProgressHandle progressHandle) {
		this.progressHandle = progressHandle;
	}
	
	@Override
	protected List<Student> doInBackground() throws Exception {
		// TODO Auto-generated method stub
        
        int pSize = progressHandle.pList.size();
        System.out.println("place size: "+pSize);
               
        Map<Integer, String>examCode = new LinkedHashMap<Integer, String>();
        examCode.put(101, "思想政治理论");
        examCode.put(199, "管理类联考综合能力");
        
        examCode.put(201, "英语一");
        examCode.put(202, "俄语");
        examCode.put(203, "日语");
        examCode.put(204, "英语二");
        
        examCode.put(301, "数学一");
        examCode.put(302, "数学二");
        examCode.put(303, "数学三");
        examCode.put(306, "西医综合");
        examCode.put(307, "中医综合");
        examCode.put(311, "教育学专业基础综合");
        examCode.put(312, "心理学专业基础综合");
        examCode.put(313, "历史学专业基础");
        
        examCode.put(408, "计算机学科专业基础综合");
        examCode.put(397, "法硕联考专业基础(法学)");
        examCode.put(398, "法硕联考专业基础(非法学)");
        examCode.put(497, "法硕联考综合(法学)");
        examCode.put(498, "法硕联考综合(非法学)");
        examCode.put(314, "数学(农)");
        examCode.put(315, "化学(农)");
        examCode.put(414, "植物生理学与生物化学");
        examCode.put(415, "动物生理学与生物化学");
        
        final int num = 60;
        int codeSize = examCode.size();
        Map<Integer, Integer>numPerClass = new LinkedHashMap<Integer, Integer>();
        Map<Integer , Integer[]> calcul = new LinkedHashMap<Integer , Integer[]>();
        for(Entry<Integer, String> entity : examCode.entrySet()){
        	
        	Integer[] values = new Integer[num+1];
        	for (int i = 0; i < values.length; i++) {
				values[i] = 0;
			}
        	calcul.put(entity.getKey(), values);
        }
        
        int index = 0;
        //循环每个人数非零的试场
        for (int i = 0; i < pSize &&  progressHandle.pList.get(i).RS > 0; i++) {
        	//当前试场每个科目的人数初始化为0
        	for(Entry<Integer, String> entity : examCode.entrySet()){
            	numPerClass.put(entity.getKey(), 0);
            }
			Place pp = progressHandle.pList.get(i);
			
			for (int j = 0; j < pp.RS; j++) {
				Student ss = progressHandle.stuList.get(index);
				int test1=0,test2=0,test3=0,test4=0;
				
				//考生ss第一单元考试代码
				try{
					test1 = Integer.valueOf(ss.ZZLLM);
					if(examCode.containsKey(test1)){
						numPerClass.put(test1, numPerClass.get(test1)+1);
					}
        			
				}catch(NumberFormatException e){}
				
				//考生ss第二单元考试代码
				try{
					test2 = Integer.valueOf(ss.WGYM);
					if(examCode.containsKey(test2)){
						numPerClass.put(test2, numPerClass.get(test2)+1);
					}
        			
				}catch(NumberFormatException e){}
				
				//考生ss第三单元考试代码
				try{
					test3 = Integer.valueOf(ss.YWK1M);
					if(examCode.containsKey(test3)){
						numPerClass.put(test3, numPerClass.get(test3)+1);
					}
        			
				}catch(NumberFormatException e){}
				
				//考生ss第四单元考试代码
				try{
					test4 = Integer.valueOf(ss.YWK2M);
					if(examCode.containsKey(test4)){
						numPerClass.put(test4, numPerClass.get(test4)+1);
					}
        			
				}catch(NumberFormatException e){}
				System.out.println(pp.SC+",编号："+(index+1)+",学生："+ss.XM+test1+":"+test2+":"+test3+":"+test4);
				index++;
			}
			
			for(Entry<Integer, String> entity : examCode.entrySet()){
				
				if(numPerClass.get(entity.getKey()) > 0){
					int nn = numPerClass.get(entity.getKey()) ;
					if(nn>=0 && nn <= num){
						calcul.get(entity.getKey())[numPerClass.get(entity.getKey())] += 1;
						System.out.println(pp.SC+" , 科目代码： "+entity.getKey()+"，需要"+numPerClass.get(entity.getKey())+"份,袋数："+calcul.get(entity.getKey())[numPerClass.get(entity.getKey())]);
					}
				}
            }
		}
        
        
        DOCWriter writer = new DOCWriter(); 
        writer.createNewDocument();    
        writer.setPageSetup(1, 20, 30, 10, 10);
        writer.setAlignment(1);
        int perPageSize = 30; //每页统计的份数大小
        //writer.setVisible(true); 
        if(size == 0){
			size = (codeSize+1)*num/perPageSize;
		}
        
        for (int i = 0; i < num/perPageSize; i++) {
        	 String str = year+"年陕西省硕士生考试统考试题使用申报表";
             writer.setFontScale("宋体", true, false,false, "0,0,0,0", 100, 20);  
             writer.insertToDocument(str);
             
             writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 15);  
             str = "考点：考试代码             考点名称(盖章)              负责人:              统计人:  ";
             writer.insertToDocument(str);
             
             List<String[]> listTab = new ArrayList<String[]>();
             String[] list = new String[perPageSize+1];
             list[0] = "科目\\需用袋数\\份数";
             for (int j = 1; j < perPageSize+1; j++) {
     			list[j] = String.valueOf(j+perPageSize*i);
     		}
             listTab.add(list);
             System.out.println(i+":1");
             for(Entry<Integer, String> code : examCode.entrySet()){
             	list = new String[perPageSize+1];
             	list[0] = code.getValue()+"("+code.getKey()+")";
             	for (int j = 1; j < perPageSize+1; j++) {
             		int value = calcul.get(code.getKey())[i*perPageSize + j];
             		if(value > 0){
             			list[j] = String.valueOf(value);
             			
             		}
          		}
             	listTab.add(list);
             }
             writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 10); 
             writer.createNewTable(codeSize+1, perPageSize+8, 1);
             for (int j = 0; j < examCode.size()+1; j++) {
     			writer.mergeCell(0, j+1, 1, j+1, 8);
     		 }
            
             for (int j = 0; j < codeSize+1; j++) {
             	writer.insertRowToTable(listTab.get(j), j);
             	lineCnt++;
             	publish(new Student());
     		}
            if( (i+1) < num/perPageSize){
          		writer.moveDown(1);
          		writer.enterDown(1);
          		writer.nextPage();
          	}
		}
        
       
        
        StringBuffer filename = new StringBuffer();
        filename.append(this.outputPath).append(File.separator).append(year).append("年试题使用申报表.doc");
        System.out.println(filename.toString());
        writer.saveAs(filename.toString());
        writer.close();
        writer.quit();
        //this.test();
		return null;
	}

	private void test(){
		int index = 0;
		int pSize = progressHandle.pList.size();
		int[] data = new int[35];
		
        for (int i = 0; i < pSize &&  progressHandle.pList.get(i).RS > 0; i++) {
        	int num = 0;
			Place pp = progressHandle.pList.get(i);
			
			for (int j = 0; j < pp.RS; j++) {
				Student ss = progressHandle.stuList.get(index);
				int test1=0;
				
				//考生ss第一单元考试代码
				try{
					test1 = Integer.valueOf(ss.YWK2M);
					if(test1 == 498){
						num++;
					}
        			
				}catch(NumberFormatException e){}
				index++;
			}
			data[num] ++;
		}
        for (int i = 0; i < data.length; i++) {
			System.out.println(i+":"+data[i]);
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
		JOptionPane.showMessageDialog(null, "成功生成试题申报表！",
				"通知", JOptionPane.INFORMATION_MESSAGE);
		progressHandle.progressBar.setValue(Values.proMinSize);
	}
	
	
}
