package com;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;
import javax.swing.SwingWorker;

public class TaskSeatdoc extends SwingWorker<List<Student>, Student>{

	private int lineCnt = 0;
	private int size = 0;
	private DefaultProgressHandle progressHandle = null;
	final private int columnSize = 4;
	private int docSize = 50;
	private String outputPath;
	private int year;
	
	public TaskSeatdoc(String outputPath, int year){
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
        int index = 0;
        System.out.println("place size: "+pSize);
        if(size == 0){
			size = pSize;
		}
        docSize = pSize;
        int docNum = (int)Math.ceil((double)pSize / docSize);
        for (int m = 0; m < docNum; m++) {
        	DOCWriter writer = new DOCWriter(); 
            writer.createNewDocument();    
            writer.setPageSetup(1, 30, 40, 20, 20);
            writer.setAlignment(1);
            //writer.setVisible(true);
            
            StringBuffer filename = new StringBuffer();
            filename.append(this.outputPath).append(File.separator).append(year).append("年考试座次表(").append(m+1).append(").doc");
            System.out.println(filename.toString());
            writer.saveAs(filename.toString());
            
            for (int i = m*docSize; i < (m+1)*docSize && i<pSize && progressHandle.pList.get(i).RS > 0 ; i++){
            	Place pp = progressHandle.pList.get(i);
            	
            	StringBuffer sb = new StringBuffer();
            	sb.append(year).append("年全国硕士生入学考试交大考点试场安排");
            	writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 18);  
                writer.insertToDocument(sb.toString());
            	
            	/*String str = "座次表";
            	writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 20);  
                writer.insertToDocument(str);*/
                 
                writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 20);  
                String str = pp.SC+":"+(int)pp.RS+"人   地点:"+pp.JS;
                writer.insertToDocument(str);
                
                int temp = index; 
                List<String[]> listTable = new ArrayList<String[]>();
                int rows = (int)Math.ceil(pp.RS/columnSize);
                int cols = columnSize;
                //循环当前教室的行数
                for (int j = 0; j < rows ; j++) {
    				String[] row = new String[cols];
    				//循环当前行的所有列
					for (int k = 0; k < cols && index < temp+pp.RS; k++) {
						//S型排列学生座位
						if( j%2 == 0){
							
							row[k] = progressHandle.stuList.get(index).KSBH+
									progressHandle.stuList.get(index).XM;
							
						}else{
							row[(int)cols - k - 1]=progressHandle.stuList.get(index).KSBH
											+progressHandle.stuList.get(index).XM;
						}
						index++;
					}
    				listTable.add(row);
    				row = null;
    			}
                
                writer.setFontScale("黑体", true, false,false, "0,0,0,0", 100, 14); 
                writer.createNewTable(rows, cols, 1);    
                writer.insertToTable(listTable); 
                if( i+1 < (m+1)*docSize && i+1 < pSize && progressHandle.pList.get(i+1).RS > 0){
                	writer.nextPage();
                }
                listTable = null;
                writer.save();
                lineCnt++;
                publish(new Student());
                System.out.println(lineCnt * 100/ size);
            }
           
            writer.save();
            writer.close();
            writer.quit();
		}
		return null;
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
		JOptionPane.showMessageDialog(null, "成功生成座次表word文档！",
				"通知", JOptionPane.INFORMATION_MESSAGE);
		progressHandle.progressBar.setValue(Values.proMinSize);
	}
	
	
}
