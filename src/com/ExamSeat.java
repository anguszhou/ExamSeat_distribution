package com;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridBagLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JToolBar;
import javax.swing.event.*;
import javax.swing.table.DefaultTableModel;

public class ExamSeat extends JFrame implements ActionListener,
ChangeListener, ListSelectionListener {

	private Map<String,List>data;
	private List<Student> stuList;
	private List<Place> pList;
	private int tabIndex , rowIndex;
	private JTabbedPane stuTab;
	private JScrollPane stuSP, seatSP;
	private JTable stuTable;
	private DefaultTableModel stuDefaultTableModel;
	private JButton importStu, matchPlace;
	private JFileChooser fileChooser;
	private JProgressBar pbar;
	DefaultProgressHandle handle = null;
	final private int docSize = 50;
	public static final int DEDAULT_WIDTH = 1000;
	public static final int DEDAULT_HEIGHT = 600;
	
	public ExamSeat() throws Exception{
		setTitle("Exam Seat Distribution");
		Toolkit kit = Toolkit.getDefaultToolkit();
		Dimension dis = kit.getScreenSize();
		setSize(DEDAULT_WIDTH, DEDAULT_HEIGHT);
		
		stuList = new ArrayList<Student>();
		pList = new ArrayList<Place>();
		data = new HashMap<String, List>();
		
		setFileChooser();
		setTab();
		setMenu();
		
		handle = new DefaultProgressHandle();
		handle.setTable(stuTable);
		handle.setProgressBar(pbar);
	}
	
	private void setFileChooser(){
		fileChooser = new JFileChooser();
		fileChooser.setFileFilter(new StuFileFilter("dbf"));
	}
	
	private void setMenu(){
		Action seatTable = new MenuAction("座次表");
		seatTable.putValue(Action.SHORT_DESCRIPTION, "生成座次表");
		seatTable.putValue(Action.MNEMONIC_KEY, new Integer('S'));
		
		Action imgTable = new MenuAction("图像表");
		imgTable.putValue(Action.SHORT_DESCRIPTION, "生成图像表");
		imgTable.putValue(Action.MNEMONIC_KEY, new Integer('I'));
		
		Action paperTable = new MenuAction("试题申报表");
		paperTable.putValue(Action.SHORT_DESCRIPTION, "生成试题申报表");
		paperTable.putValue(Action.MNEMONIC_KEY, new Integer('P'));
		
		JToolBar toolBar = new JToolBar();
		toolBar.add(seatTable);
		toolBar.add(imgTable);
		toolBar.addSeparator();
		toolBar.add(paperTable);
		add(toolBar,BorderLayout.NORTH);
	}
	
	class MenuAction extends AbstractAction{
		public MenuAction(String name){
			super(name);
		}

		@Override
		public void actionPerformed(ActionEvent e) {
			// TODO Auto-generated method stub
			if(e.getActionCommand().equals("座次表")){
				int select = JOptionPane.showConfirmDialog(null, "确定开始生成座次表？",
											"警告",JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE);
				if(select == JOptionPane.YES_OPTION){
					handle.progressBar.setValue(Values.proMinSize);
					if(handle.stuList != null &&!handle.stuList.isEmpty() &&
						handle.pList!=null && !handle.pList.isEmpty()	 ){
						
						JFileChooser fc2 = new JFileChooser();
						fc2.setFileFilter(new StuFileFilter("dir"));
						fc2.setDialogTitle("请选择导出座次表的路径");
						fc2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
						//fc.setMultiSelectionEnabled(true);
						
						int result2 = fc2.showOpenDialog(ExamSeat.this);
						File output = null;
						if(result2 == JFileChooser.APPROVE_OPTION){
							output = fc2.getSelectedFile();
							TaskSeatdoc task = new TaskSeatdoc(output.getAbsolutePath());
							task.setProgressHandle(handle);
							task.execute();
							task = null;
						}else if(result2 == JFileChooser.CANCEL_OPTION){
							return;
						}
						
					}else{
						JOptionPane.showMessageDialog(null, "请先导入考生和试场数据",
								"警告",JOptionPane.INFORMATION_MESSAGE);
					}
					
				}
			}else if(e.getActionCommand().equals("图像表")){
				int select = JOptionPane.showConfirmDialog(null, "确定开始生成图像表？",
						"警告",JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE);
				if(select == JOptionPane.YES_OPTION){
					handle.progressBar.setValue(Values.proMinSize);
					if(handle.stuList != null &&!handle.stuList.isEmpty()&&
							handle.pList!=null && !handle.pList.isEmpty()	 ){
						
						JFileChooser fc = new JFileChooser();
						fc.setFileFilter(new StuFileFilter("dir"));
						fc.setDialogTitle("请导入考生照片数据");
						fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
						//fc.setMultiSelectionEnabled(true);
						
						int result = fc.showOpenDialog(ExamSeat.this);
						File file = null;
						if(result == JFileChooser.APPROVE_OPTION){
							file = fc.getSelectedFile();
							String prefix = new String();
							String suffix = new String();
							if( file.isDirectory()){
								File[] files = file.listFiles();
								for(File tmp : files){
									if(tmp.isFile()){
										String name = tmp.getName();
										prefix = name.substring(0, name.lastIndexOf('_')+1);
										suffix = name.substring(name.lastIndexOf('.'));
										break;
									}
								}
								System.out.println(prefix+"::"+suffix);
							}else{
								JOptionPane.showMessageDialog(null, "请选择包含考生照片的文件夹！",
										"警告",JOptionPane.INFORMATION_MESSAGE);
								return ;
							}
							try{
								JFileChooser fc2 = new JFileChooser();
								fc2.setFileFilter(new StuFileFilter("dir"));
								fc2.setDialogTitle("请选择导出图像表的路径");
								fc2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
								//fc.setMultiSelectionEnabled(true);
								
								StringBuffer sb = new StringBuffer();
								sb.append(file.getAbsolutePath()).append(File.separator).append(prefix);
								
								int result2 = fc2.showOpenDialog(ExamSeat.this);
								File output = null;
								if(result2 == JFileChooser.APPROVE_OPTION){
									output = fc2.getSelectedFile();
									int notZeroNum = getNotZeroClass(handle.pList);
									handle.size = notZeroNum;
									handle.lineCnt = 0;
									int docNum = (int)Math.ceil((double)handle.pList.size() / Values.docSize);
									for (int i = 0; i < docNum; i++) {
										TaskImgdoc task = new TaskImgdoc(sb.toString(),suffix,output.getAbsolutePath(),i);
										task.setProgressHandle(handle);
										task.execute();
									}
									
								}else if(result == JFileChooser.CANCEL_OPTION){
									return;
								}
								
							}catch(Exception e3){
								e3.printStackTrace();
							}	
						}else if(result == JFileChooser.CANCEL_OPTION){
							return;
						}
						
						
					}else{
						JOptionPane.showMessageDialog(null, "请先导入考生和试场数据",
								"警告",JOptionPane.INFORMATION_MESSAGE);
					}
				
				}
			}else if(e.getActionCommand().equals("试题申报表")){
				int select = JOptionPane.showConfirmDialog(null, "确定开始生成试题申报表？",
						"警告",JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE);
				if(select == JOptionPane.YES_OPTION){
					handle.progressBar.setValue(Values.proMinSize);
					if(handle.stuList != null &&!handle.stuList.isEmpty()&&
							handle.pList!=null && !handle.pList.isEmpty()	 ){
						
						try{
							JFileChooser fc = new JFileChooser();
							fc.setFileFilter(new StuFileFilter("dir"));
							fc.setDialogTitle("请选择导出试题申报表的路径");
							fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
							//fc.setMultiSelectionEnabled(true);
							
							int result = fc.showOpenDialog(ExamSeat.this);
							File output = null;
							if(result == JFileChooser.APPROVE_OPTION){
								output = fc.getSelectedFile();
								
								TaskCalPaperdoc task = new TaskCalPaperdoc(output.getAbsolutePath());
								task.setProgressHandle(handle);
								task.execute();
								
							}else if(result == JFileChooser.CANCEL_OPTION){
								return;
							}
							
						}catch(Exception e3){
							e3.printStackTrace();
						}	
						
						
					}else{
						JOptionPane.showMessageDialog(null, "请先导入考生和试场数据",
								"警告",JOptionPane.INFORMATION_MESSAGE);
					}
				
				}
			}else{
				JOptionPane.showMessageDialog(null, "你点击了："+e.getActionCommand().toString(),
											"通知",JOptionPane.INFORMATION_MESSAGE);
			}
		}
		
	}
	
	public static int getNotZeroClass(List<Place> pList){
		int num = 0;
		for (int i = 0; i < pList.size(); i++) {
			Place pp = pList.get(i);
			if(pp.RS > 0){
				num++;
			}
		}
		return num;
	}
	
	private void setTab(){
		tabIndex = 0 ;
		setStuSP();
		stuTab = new JTabbedPane();
		stuTab.addTab("考生信息", stuSP);
		//stuTab.addTab("111", seatSP);
		
		stuTab.addChangeListener(this);
	}
	
	private void setStuSP(){
		rowIndex = -1;
		JPanel input = new JPanel();
		input.setLayout(new GridBagLayout());
		
		importStu = new JButton("导入考生信息");
		importStu.setMnemonic(KeyEvent.VK_I);
		importStu.addActionListener(this);
		matchPlace = new JButton("分配考场");
		matchPlace.setMnemonic(KeyEvent.VK_M);
		matchPlace.addActionListener(this);
		pbar = new JProgressBar(Values.proMinSize,Values.proMaxSize);
		JLabel pName = new JLabel("进度：");
		
		input.add(importStu, new GBC(0,0).setAnchor(GBC.EAST));
		input.add(matchPlace, new GBC(1,0).setAnchor(GBC.WEST));
		input.add(pName, new GBC(0,1).setAnchor(GBC.EAST));
		input.add(pbar, new GBC(1,1,5,1).setFill(GBC.HORIZONTAL).setWeight(100, 0).setInsets(1));
		
		setStuTable();
		JScrollPane tableScroll = new JScrollPane(stuTable);
		
		JPanel main = new JPanel();
		main.setLayout(new BorderLayout());
		main.add(input, BorderLayout.NORTH);
		main.add(tableScroll, BorderLayout.CENTER);
		
		stuSP = new JScrollPane();
		stuSP.setViewportView(main);
		add(main);
	}
	
	private void setStuTable(){
		String[] name = {"编号","考生编号","姓名","试场","教室","政治理论","外国语","业务课1","业务课2"};
		Object[][] data = null;
		stuDefaultTableModel = new DefaultTableModel(data, name){
			public boolean isCellEditable(int row, int column){
				return false;
			}
		};
		
		stuTable = new JTable(stuDefaultTableModel);
		stuTable.getSelectionModel().addListSelectionListener(this);
		
	}
	
	@Override
	public void valueChanged(ListSelectionEvent e) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void stateChanged(ChangeEvent e) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		if(e.getActionCommand().equals("导入考生信息")){
			fileChooser.setDialogTitle("导入考生信息");
			fileChooser.setSelectedFile(new File("*.dbf"));
			int result = fileChooser.showOpenDialog(this);
			File file = null;
			if(result == JFileChooser.APPROVE_OPTION){
				file = fileChooser.getSelectedFile();
				if(!file.exists()){
					JOptionPane.showMessageDialog(null, "文件不存在",
							"警告", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				try{
					stuList.clear();
					TaskReadDBF task = new TaskReadDBF(file.getAbsolutePath());
					task.setProgressHandle(handle);
					task.execute();
					task = null;
					
				}catch(Exception e3){
					e3.printStackTrace();
				}		
				
			}else if(result == JFileChooser.CANCEL_OPTION){
				return;
			}
		}else if(e.getActionCommand().equals("分配考场")){
			fileChooser.setDialogTitle("请导入考场信息数据");
			fileChooser.setSelectedFile(new File("*.dbf"));
			int result = fileChooser.showOpenDialog(this);
			File file = null;
			if(result == JFileChooser.APPROVE_OPTION){
				file = fileChooser.getSelectedFile();
				if(!file.exists()){
					JOptionPane.showMessageDialog(null, "文件不存在",
							"警告", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
				try{
					TaskDistribute task = new TaskDistribute(file.getAbsolutePath());
					task.setProgressHandle(handle);
					task.execute();
					task = null;
					
				}catch(Exception e3){
					e3.printStackTrace();
				}	
			}else if(result == JFileChooser.CANCEL_OPTION){
				return;
			}
				
		}
	}

	public static void main(String[] arvg){
		java.awt.EventQueue.invokeLater(new Runnable() {
			public void run() {
				ExamSeat es = null;
				try{
					es = new ExamSeat();
				}catch(Exception e){
					e.printStackTrace();
				}
				es.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
				es.setVisible(true);
			}
		});
	}
}
