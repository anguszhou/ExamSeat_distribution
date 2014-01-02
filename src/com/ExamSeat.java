package com;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyEvent;
import java.io.File;
import java.util.Calendar;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.BorderFactory;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JToolBar;
import javax.swing.SwingConstants;
import javax.swing.border.TitledBorder;
import javax.swing.event.*;
import javax.swing.table.DefaultTableModel;

public class ExamSeat extends JFrame implements ActionListener,
ChangeListener, ListSelectionListener,ItemListener {

	private Map<String,List>data;
	private List<Student> stuList;
	private List<Place> pList;
	private int tabIndex , rowIndex;
	private JTabbedPane stuTab;
	private JScrollPane stuSP, teacherSP;
	private JTable stuTable , disedPlaceTable , nDisTeaTable;
	private DefaultTableModel stuDefaultTableModel , disedPlaceTableModel, nDisTeaTableModel;
	private JButton importStu, matchPlace;
	private JFileChooser fileChooser;
	private JProgressBar pbar , teacherBar;
	private JComboBox placeBox;
	private DefaultComboBoxModel placeBoxMode;
	private JTextField curPlaceField, curTeaField;
	DefaultProgressHandle handle = null;
	final private int docSize = 50;
	public static final int DEDAULT_WIDTH = 1000;
	public static final int DEDAULT_HEIGHT = 600;
	private int year = 0;
	private int tNum = 2;
	private int MBANum = 4;
	private String marks = "1.核对照片(本人、准考证)2.考生签字(缺考由监考老师在签字处注'缺考')" +
			"3.监考人员在右下角签字4.本表与考场记录一并交回。" , type = "";
	private List<Teacher> curTea = new ArrayList<Teacher>();
	
	public ExamSeat() throws Exception{
		setTitle("Exam Seat Distribution");
		Toolkit kit = Toolkit.getDefaultToolkit();
		Dimension dis = kit.getScreenSize();
		setSize(DEDAULT_WIDTH, DEDAULT_HEIGHT);
		
		stuList = new ArrayList<Student>();
		pList = new ArrayList<Place>();
		data = new HashMap<String, List>();
		year = Calendar.getInstance().get(Calendar.YEAR);
		
		setFileChooser();
		setTab();
		setMenu();
		
		handle = new DefaultProgressHandle();
		handle.setTable(stuTable , disedPlaceTable , nDisTeaTable);
		handle.setProgressBar(pbar, teacherBar);
		handle.setCombBox(placeBox);
	}
	
	private void setFileChooser(){
		fileChooser = new JFileChooser();
		fileChooser.setFileFilter(new StuFileFilter("dbf"));
	}
	
	private void setMenu(){
		
		Action setPara = new MenuAction("参数");
		setPara.putValue(Action.SHORT_DESCRIPTION, "设置年份等");
		
		Action seatTable = new MenuAction("座次表");
		seatTable.putValue(Action.SHORT_DESCRIPTION, "生成座次表");
		seatTable.putValue(Action.MNEMONIC_KEY, new Integer('S'));
		
		Action imgTable = new MenuAction("图像表");
		imgTable.putValue(Action.SHORT_DESCRIPTION, "生成图像表");
		imgTable.putValue(Action.MNEMONIC_KEY, new Integer('I'));
		
		Action paperTable = new MenuAction("试题申报表");
		paperTable.putValue(Action.SHORT_DESCRIPTION, "生成试题申报表");
		paperTable.putValue(Action.MNEMONIC_KEY, new Integer('P'));
		
		Action teaTable = new MenuAction("监考表");
		teaTable.putValue(Action.SHORT_DESCRIPTION, "生成监考按排表");
		teaTable.putValue(Action.MNEMONIC_KEY, new Integer('T'));
		
		JToolBar toolBar = new JToolBar();
		toolBar.add(setPara);
		toolBar.addSeparator();
		toolBar.add(seatTable);
		toolBar.add(imgTable);
		toolBar.addSeparator();
		toolBar.add(paperTable);
		toolBar.add(teaTable);
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
							TaskSeatdoc task = new TaskSeatdoc(output.getAbsolutePath(), year);
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
									handle.size = notZeroNum *2;
									handle.lineCnt = 0;
									int docNum = (int)Math.ceil((double)handle.pList.size() / Values.docSize);
									for (int i = 0; i < docNum; i++) {
										TaskImgdoc task = new TaskImgdoc(sb.toString(),suffix,output.getAbsolutePath(),i,marks, year);
										task.setProgressHandle(handle);
										task.execute();
									}
									/*TaskImgdoc task = new TaskImgdoc(sb.toString(),suffix,output.getAbsolutePath(),4,marks, year);
									task.setProgressHandle(handle);
									task.execute();*/
									
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
								
								TaskCalPaperdoc task = new TaskCalPaperdoc(output.getAbsolutePath() , year);
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
			}else if(e.getActionCommand().equals("参数")){
				new InputDialog(ExamSeat.this);
				
			}else if(e.getActionCommand().equals("监考表")){
				int select = JOptionPane.showConfirmDialog(null, "确定开始生成监考按排表？",
						"警告",JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE);
				if(select == JOptionPane.YES_OPTION){
					handle.teacherBar.setValue(Values.proMinSize);
					if( //handle.pTeaList != null &&!handle.pTeaList.isEmpty()&&
							handle.teaList!=null && !handle.teaList.isEmpty()	 ){
						
						JFileChooser fc = new JFileChooser();
						fc.setFileFilter(new StuFileFilter("dir"));
						fc.setDialogTitle("请导入"+type+"照片数据");
						fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
						//fc.setMultiSelectionEnabled(true);
						
						int result = fc.showOpenDialog(ExamSeat.this);
						File file = null;
						if(result == JFileChooser.APPROVE_OPTION){
							file = fc.getSelectedFile();
							try{
								JFileChooser fc2 = new JFileChooser();
								fc2.setFileFilter(new StuFileFilter("dir"));
								fc2.setDialogTitle("请选择导出监考按排表的路径");
								fc2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
								//fc.setMultiSelectionEnabled(true);
								
								int result2 = fc2.showOpenDialog(ExamSeat.this);
								File output = null;
								if(result2 == JFileChooser.APPROVE_OPTION){
									output = fc2.getSelectedFile();
									TaskTeadoc task = new TaskTeadoc(file.getAbsolutePath(),output.getAbsolutePath(),year,type);
									task.setProgressHandle(handle);
									task.execute();
									
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
						JOptionPane.showMessageDialog(null, "请先导入监考老师或工作人员数据",
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
		setTeaSP();
		stuTab = new JTabbedPane();
		stuTab.addTab("分配考场", stuSP);
		stuTab.addTab("分配监考老师", teacherSP);
		
		stuTab.addChangeListener(this);
		add(stuTab);
	}
	
	private void setTeaSP(){
		teacherSP = new JScrollPane();
		JPanel main = new JPanel();
		main.setLayout(new BorderLayout());
		
		JPanel jp = new JPanel();
		jp.setLayout(new GridBagLayout());
		
		//导入按钮和进度条
		JButton inputPlace = new JButton("导入考场信息");
		JButton inputTea = new JButton("导入监考老师");
		JButton inputWork = new JButton("导入工作人员");
		JButton export = new JButton("导出监考信息");
		inputPlace.addActionListener(this);
		inputTea.addActionListener(this);
		inputWork.addActionListener(this);
		export.addActionListener(this);
		teacherBar = new JProgressBar(Values.proMinSize,Values.proMaxSize);
		jp.add(inputPlace, new GBC(0,0,2,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		jp.add(inputTea, new GBC(2,0,2,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		jp.add(inputWork, new GBC(4,0,1,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		jp.add(export, new GBC(6,0,2,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		jp.add(teacherBar, new GBC(0,1,8,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(1, 0).setInsets(5).setIpad(0, 0));
		//未分配监考老师的试场 下拉框
		placeBoxMode = new DefaultComboBoxModel();
		placeBox = new JComboBox(placeBoxMode);
		placeBox.addItemListener(this);
		
		JLabel placename = new JLabel("未分配的考场：");
		JButton distribute = new JButton("分配");
		JButton reset = new JButton("重置");
		distribute.addActionListener(this);
		reset.addActionListener(this);
		jp.add(placename, new GBC(0,2,2,1).setAnchor(GBC.EAST)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		jp.add(placeBox, new GBC(2,2,4,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(1, 0).setInsets(5).setIpad(0, 0));
		jp.add(distribute, new GBC(6,2,1,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		jp.add(reset, new GBC(7,2,1,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		
		//当前选中的考场和监考老师
		JLabel curDistribute = new JLabel("正在分配的考场：");
		curPlaceField = new JTextField();
		curTeaField = new JTextField();
		curPlaceField.setEditable(false);
		curTeaField.setEditable(false);
		
		JButton clear = new JButton("清除");
		clear.addActionListener(this);
		jp.add(curDistribute, new GBC(0,3,2,1).setAnchor(GBC.EAST)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		jp.add(curPlaceField, new GBC(2,3,2,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(0.2, 0).setInsets(5).setIpad(0, 0));
		jp.add(curTeaField, new GBC(4,3,2,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(1, 0).setInsets(5).setIpad(0, 0));
		jp.add(clear, new GBC(6,3,1,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		
		jp.setBorder(BorderFactory.createTitledBorder(BorderFactory.createLineBorder(Color.BLUE,2),
	                "",TitledBorder.CENTER,TitledBorder.TOP));
		
		setPlaceAndTeaTable();
		//已分配的考场
		JScrollPane jp2 = new JScrollPane(disedPlaceTable);
		jp2.setBorder(BorderFactory.createTitledBorder(BorderFactory.createLineBorder(Color.BLUE,2),
                "已分配的考场",TitledBorder.CENTER,TitledBorder.TOP));
		
		//未分配监考老师
		JScrollPane jp3 = new JScrollPane(nDisTeaTable);
		jp3.setBorder(BorderFactory.createTitledBorder(BorderFactory.createLineBorder(Color.BLUE,2),
                "未分配的监考老师",TitledBorder.CENTER,TitledBorder.TOP));
		
		main.add(jp , BorderLayout.NORTH);
		main.add(jp2 , BorderLayout.CENTER);
		main.add(jp3 , BorderLayout.SOUTH);
		teacherSP.setViewportView(main);
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
	
	private void setPlaceAndTeaTable(){
		String[] palce = {"编号","试场","地点","人数","考试类型","监考老师"};
		Object[][] data = null;
		disedPlaceTableModel= new DefaultTableModel(data, palce){
			public boolean isCellEditable(int row, int column){
				return false;
			}
		};
		
		disedPlaceTable = new JTable(disedPlaceTableModel);
		disedPlaceTable.getSelectionModel().addListSelectionListener(this);
		disedPlaceTable.setPreferredScrollableViewportSize(new Dimension(900, 150));
		
		String[] tea = {"编号","监考编号","姓名","性别","职务","单位","是否"};
		Object[][] data2 = null;
		nDisTeaTableModel= new DefaultTableModel(data2, tea){
			public boolean isCellEditable(int row, int column){
				return false;
			}
		};
		
		nDisTeaTable = new JTable(nDisTeaTableModel);
		nDisTeaTable.getSelectionModel().addListSelectionListener(this);
		nDisTeaTable.setPreferredScrollableViewportSize(new Dimension(900, 150));
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
		
		input.add(importStu, new GBC(0,0 ,2,1).setAnchor(GBC.EAST).setFill(GBC.BOTH)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		input.add(matchPlace, new GBC(2,0,2,1).setAnchor(GBC.WEST).setFill(GBC.BOTH)
		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
		input.add(pbar, new GBC(0,1,8,1).setAnchor(GBC.CENTER).setFill(GBC.BOTH)
		 		.setWeight(1, 0).setInsets(5).setIpad(0, 0));
		input.setBorder(BorderFactory.createTitledBorder(BorderFactory.createLineBorder(Color.BLUE,2),
                "",TitledBorder.CENTER,TitledBorder.TOP));
		
		setStuTable();
		JScrollPane tableScroll = new JScrollPane(stuTable);
		tableScroll.setBorder(BorderFactory.createTitledBorder(BorderFactory.createLineBorder(Color.BLUE,2),
                "考生信息",TitledBorder.CENTER,TitledBorder.TOP));
		
		JPanel main = new JPanel();
		main.setLayout(new BorderLayout());
		main.add(input, BorderLayout.NORTH);
		main.add(tableScroll, BorderLayout.CENTER);
		
		stuSP = new JScrollPane();
		stuSP.setViewportView(main);
	}
	
	/*
	*当用鼠标对表格进行选取，在响应行选取变化事件（ListSelectionListener）时，
	*鼠标按下会响应一次，鼠标释放又会响应一次，因此一次鼠标的点击会有两次事件响应（按下和释放）。
	*前者的事件属性中getValueIsAdjusting()=true，后者是false。因此，可以通过判断
	*getValueIsAdjusting()来区别鼠标按下和释放，进行不同的操作。而用键盘的上下键选取时，只有一次事件响应。
	*/
	@Override
	public void valueChanged(ListSelectionEvent e) {
		// TODO Auto-generated method stub
		if(e.getSource().equals(disedPlaceTable.getSelectionModel())){
			int[] rows = disedPlaceTable.getSelectedRows();
			if(rows.length == 1 && e.getValueIsAdjusting() == true){
				rowIndex = rows[0];
				StringBuffer sb = new StringBuffer();
				sb.append("[").append(disedPlaceTable.getValueAt(rowIndex, 0)).append("]")
								.append(disedPlaceTable.getValueAt(rowIndex, 1))
								.append("[人数:").append(disedPlaceTable.getValueAt(rowIndex, 3)).append("]")
								.append("[类型:").append(disedPlaceTable.getValueAt(rowIndex, 4)).append("]");
				curPlaceField.setText(sb.toString());
				curTeaField.setText(disedPlaceTable.getValueAt(rowIndex, 5).toString());
			}
			
		}else if(e.getSource().equals(nDisTeaTable.getSelectionModel())){
			int[] rows = nDisTeaTable.getSelectedRows();
			if(rows.length == 1 && e.getValueIsAdjusting() == true){
				rowIndex = rows[0];
				System.out.println(rowIndex+":"+nDisTeaTable.getValueAt(rowIndex, 0).toString()+","
						+nDisTeaTable.getValueAt(rowIndex, 1).toString()+","
						+nDisTeaTable.getValueAt(rowIndex, 2).toString());
				
				int index = Integer.valueOf(nDisTeaTable.getValueAt(rowIndex, 0).toString());
				for (int i = 0; i < curTea.size(); i++) {
					if(curTea.get(i).id == index-1){
						return;
					}
				}
				
				StringBuffer sb = new StringBuffer();
				sb.append(curTeaField.getText()).append(" [").append(nDisTeaTable.getValueAt(rowIndex, 0).toString())
					.append(":").append(nDisTeaTable.getValueAt(rowIndex, 2).toString()).append("]");
				curTeaField.setText(sb.toString());
				
				curTea.add(handle.teaList.get(index - 1));
			}
			
		}
	}

	@Override
	public void stateChanged(ChangeEvent e) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void itemStateChanged(ItemEvent e) {
		// TODO Auto-generated method stub
		if (e.getStateChange()==ItemEvent.SELECTED){//当用户的选择改变时.
			curPlaceField.setText(e.getItem().toString());
			curTeaField.setText("");
			curTea.clear();
			nDisTeaTable.clearSelection();
			disedPlaceTable.clearSelection();
	   	  }
	}
	
	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		if(e.getActionCommand().equals("导入考生信息")){
			if(handle.stuList != null ){
				JOptionPane.showMessageDialog(null, "已经导入了考生信息，无法继续导入！如需重新导入数据，请重新运行本程序。",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
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
			if(handle.pList != null){
				JOptionPane.showMessageDialog(null, "已经为考生分配了考场，无法继续分配！如需重新分配，请重新运行本程序。",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			if(handle.stuList == null || handle.stuList.isEmpty() ){
				JOptionPane.showMessageDialog(null, "请先导入考生信息！",
						"警告",JOptionPane.INFORMATION_MESSAGE);
				return ;
			}else{
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
		}else if(e.getActionCommand().equals("导入监考老师")){
			if(handle.teaList != null){
				JOptionPane.showMessageDialog(null, "已经导入了监考老师或工作人员，无法继续导入！如需重新导入数据，请重新运行本程序。",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			fileChooser.setDialogTitle("请导入监考老师的数据");
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
					type = "监考人员";
					TaskReadTeaDBF task = new TaskReadTeaDBF(file.getAbsolutePath());
					task.setProgressHandle(handle);
					task.execute();
					task = null;
					
				}catch(Exception e3){
					e3.printStackTrace();
				}	
			}else if(result == JFileChooser.CANCEL_OPTION){
				return;
			}
			
		}else if(e.getActionCommand().equals("导入工作人员")){
			if(handle.teaList != null){
				JOptionPane.showMessageDialog(null, "已经导入了监考老师或工作人员，无法继续导入！如需重新导入数据，请重新运行本程序。",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			fileChooser.setDialogTitle("请导入工作人员的数据");
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
					type = "工作人员";
					TaskReadTeaDBF task = new TaskReadTeaDBF(file.getAbsolutePath());
					task.setProgressHandle(handle);
					task.execute();
					task = null;
					
				}catch(Exception e3){
					e3.printStackTrace();
				}	
			}else if(result == JFileChooser.CANCEL_OPTION){
				return;
			}
			
		}else if(e.getActionCommand().equals("导入考场信息")){
			if(handle.pTeaList != null){
				JOptionPane.showMessageDialog(null, "已经导入了考场信息，无法继续导入！",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			fileChooser.setDialogTitle("请导入考场的数据");
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
					TaskPlaceDBF task = new TaskPlaceDBF(file.getAbsolutePath());
					task.setProgressHandle(handle);
					task.execute();
					task = null;
					
				}catch(Exception e3){
					e3.printStackTrace();
				}	
			}else if(result == JFileChooser.CANCEL_OPTION){
				return;
			}
		}else if(e.getActionCommand().equals("导出监考信息")){
			if(handle.teaList==null || handle.teaList.isEmpty()
					|| handle.pTeaList==null || handle.pTeaList.isEmpty()){
				JOptionPane.showMessageDialog(null, "没有监考老师或考场信息，无法导出！",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			/*
			if(placeBoxMode.getSize() > 0){
				JOptionPane.showMessageDialog(null, "考场还未分配完毕，无法导出！",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			*/
			try{
				JFileChooser fc = new JFileChooser();
				fc.setFileFilter(new StuFileFilter("dir"));
				fc.setDialogTitle("请选择导出监考数据的路径");
				fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				//fc.setMultiSelectionEnabled(true);
				
				int result = fc.showOpenDialog(ExamSeat.this);
				File output = null;
				if(result == JFileChooser.APPROVE_OPTION){
					output = fc.getSelectedFile();
					
					TaskExportTea task = new TaskExportTea(output.getAbsolutePath());
					task.setProgressHandle(handle);
					task.year = this.year;
					task.execute();
					
				}else if(result == JFileChooser.CANCEL_OPTION){
					return;
				}
				
			}catch(Exception e3){
				e3.printStackTrace();
			}	
			
			
		}else if(e.getActionCommand().equals("分配")){
			
			if(curPlaceField.getText().equals("") || placeBoxMode.getSize() == 0){
				JOptionPane.showMessageDialog(null, "当前没有要分配的考场！",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}else if(curTea.isEmpty()){
				JOptionPane.showMessageDialog(null, "请先选择要分配的监考老师！",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			//判断当前考场需要分配的监考老师人数是否超过限制
			String curPlace = curPlaceField.getText();
			String pid =  curPlace.substring(curPlace.indexOf("[")+1, curPlace.indexOf("]"));
			int placeIndex = Integer.valueOf(pid);
			
			//排除已分配的情况
			for (int i = 0; i < disedPlaceTableModel.getRowCount(); i++) {
				if(Integer.valueOf(disedPlaceTableModel.getValueAt(i, 0).toString()) == placeIndex){
					JOptionPane.showMessageDialog(null, "当前考场已经分配了监考老师！",
							"警告", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
			}
			
			Place pp = handle.pTeaList.get(placeIndex-1);
			if(pp.BZ.toUpperCase().startsWith("M")){
				if(curTea.size() > MBANum ){
					JOptionPane.showMessageDialog(null, "MBA考场监考人数应为:"+MBANum+",当前选择的监考人数为:"+curTea.size()+",超过限制,请重新监考老师！",
							"警告", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
			}else{
				if(curTea.size() >  tNum){
					JOptionPane.showMessageDialog(null, "该考场监考人数应为:"+tNum+",当前选择的监考人数为:"+curTea.size()+",超过限制,请重新监考老师！",
							"警告", JOptionPane.INFORMATION_MESSAGE);
					return;
				}
			}
			//1设置监考老师的考场信息
			for (Teacher t : curTea) {
				t.SC = pp.SC;
				t.DD = pp.JS;
				System.out.println("curTea size:"+curTea.size()+":"+t.id+":"+t.XM);
			}
			//2按编号在已分配考场列表中加入该考场
			StringBuffer sb = new StringBuffer();
			for (int i = 0; i < curTea.size(); i++) {
				sb.append("[").append(curTea.get(i).id+1).append(":").append(curTea.get(i).XM).append("] ");
			}
			Object[] newRow = {pp.id+1 , pp.SC , pp.JS, (int)pp.RS ,pp.BZ ,sb.toString()};
			int rowNum = 0;
			for (int i = 0; i < disedPlaceTableModel.getRowCount(); i++) {
				if(pp.id+1 > Integer.valueOf(disedPlaceTableModel.getValueAt(i, 0).toString())){
					rowNum++;
				}
			}
			disedPlaceTableModel.insertRow(rowNum, newRow);
			//3删除未分配考场列表中的该考场和未分配监考老师列表中的相关老师
			//nDisTeaTableModel删除行时，下一次getRowCount会变化！
			for (int i = 0; i < curTea.size(); i++) {
				for (int j = 0; j < nDisTeaTableModel.getRowCount(); j++) {
					if( curTea.get(i).id+1 == Integer.valueOf(nDisTeaTable.getValueAt(j, 0).toString()) ){
						nDisTeaTableModel.removeRow(j);
						break;
					}
				}
			}
			/*删除行时会触发combox的itemStateChanged事件，会清空curTea列表中的数据，所以放在最后处理
			*但是删除最后一行时不会触发itemStateChanged事件！
			*/
			for (int i = 0; i < placeBoxMode.getSize(); i++) {
				if(placeBoxMode.getElementAt(i).toString().equals(curPlace)){
					placeBoxMode.removeElementAt(i);
					break;
				}
			}
			//4清除正在分配的输入框内容和curTea内容
			if(placeBoxMode.getSize() == 0){
				curPlaceField.setText("");
			}
			curTeaField.setText("");
			curTea.clear();
			
		}else if(e.getActionCommand().equals("重置")){
			if(curPlaceField.getText().equals("") || curTeaField.getText().equals("")){
				JOptionPane.showMessageDialog(null, "请选择要重置的考场！",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			String curPlace = curPlaceField.getText();
			String pid =  curPlace.substring(curPlace.indexOf("[")+1, curPlace.indexOf("]"));
			int placeIndex = Integer.valueOf(pid);
			
			boolean flag = false , flag2 = false;
			for (int i = 0; i < disedPlaceTableModel.getRowCount(); i++) {
				if(Integer.valueOf(disedPlaceTableModel.getValueAt(i, 0).toString()) == placeIndex){
					flag = true;
					if(curTeaField.getText().toString().equals(disedPlaceTableModel.getValueAt(i, 5).toString())){
						flag2 = true;
					}
					break;
				}
			}
			if(!flag){
				JOptionPane.showMessageDialog(null, "该考场还未分配监考老师，无法重置，请重新选择！",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}else if(!flag2){
				JOptionPane.showMessageDialog(null, "该考场的监考老师与已分配的不符，请清除后重新选择！",
						"警告", JOptionPane.INFORMATION_MESSAGE);
				return;
			}
			//1.重置监考老师的数据和监考人员加回未分配人员列表
			String[] curTeas = curTeaField.getText().split("]");
			
			for (int i = 0; i < curTeas.length; i++) {
				//System.out.println(i+":"+curTeas[i]+":"+curTeas[i].length());
				if(curTeas[i].length() < 3){
					break;
				}
				int teaIndex = Integer.valueOf(curTeas[i].substring(curTeas[i].indexOf("[")+1, curTeas[i].indexOf(":")));
				
				System.out.println("重置前："+teaIndex+":"+handle.teaList.get(teaIndex-1).XM+":"+handle.teaList.get(teaIndex-1).SC);
				Teacher tt = handle.teaList.get(teaIndex-1);
				tt.SC = "";
				tt.DD = "";
				System.out.println("重置后："+teaIndex+":"+handle.teaList.get(teaIndex-1).XM+":"+handle.teaList.get(teaIndex-1).SC);
			
				int rowNum = 0;
				for (int j = 0; j < nDisTeaTableModel.getRowCount(); j++) {
					if( teaIndex > Integer.valueOf(nDisTeaTableModel.getValueAt(j, 0).toString()) ){
						rowNum++;
					}else{
						break;
					}
				}
				Object[] newRow = {tt.id+1 , tt.KH, tt.XM, tt.XB ,tt.ZW ,tt.DW, tt.SF};
				nDisTeaTableModel.insertRow(rowNum, newRow);
			}
			//2.将考场加回combox
			int rowNum = 0;
			for (int i = 0; i < placeBoxMode.getSize(); i++) {
				String  text = placeBoxMode.getElementAt(i).toString();
				int tmpIndex = Integer.valueOf(text.substring(text.indexOf("[")+1, text.indexOf("]")));
				if( placeIndex > tmpIndex){
					rowNum++;
				}else{
					break;
				}
			}
			placeBoxMode.insertElementAt(curPlace, rowNum);
			placeBoxMode.setSelectedItem(curPlace);
			
			//3.删除已分配考场列表中该条数据
			for (int i = 0; i < disedPlaceTableModel.getRowCount(); i++) {
				if(Integer.valueOf(disedPlaceTableModel.getValueAt(i, 0).toString()) == placeIndex){
					disedPlaceTableModel.removeRow(i);
					break;
				}
			}
			//4清除正在分配的输入框内容和curTea内容
			curTeaField.setText("");
			curTea.clear();
			
		}else if(e.getActionCommand().equals("清除")){
			if(placeBoxMode.getSize() > 0){
				curPlaceField.setText(placeBoxMode.getSelectedItem().toString());
			}
			curTeaField.setText("");
			curTea.clear();
			nDisTeaTable.clearSelection();
			disedPlaceTable.clearSelection();
		}
	}

	class InputDialog implements ActionListener{
	      JTextField   marksField;
	      JComboBox yearField , tNumField , MBANumField;
	      JDialog dialog;
	      int[] teacherNum , yearNum;
	      
	      public void actionPerformed(ActionEvent e){
	      	 String cmd=e.getActionCommand();
	      	 if (cmd.equals("确定")){
	      		 year = yearNum[yearField.getSelectedIndex()];
	      		 tNum = teacherNum[tNumField.getSelectedIndex()];
	      		 MBANum = teacherNum[MBANumField.getSelectedIndex()];
	      		 marks = marksField.getText();
	      		 System.out.println(year+","+tNum+","+MBANum+","+marks);
	      		 dialog.dispose();
	      	 	
	      	 }else if (cmd.equals("取消")){
	      	   dialog.dispose();	
	      	 }
	      }
	      InputDialog(JFrame f){
	    	
	    	teacherNum = new int[10];
	    	yearNum = new int[6];
	    	for (int i = 0; i < teacherNum.length; i++) {
				teacherNum[i] = i+1;
			}
	    	for (int i = 0; i < yearNum.length; i++) {
	    		int offset = -2;
				yearNum[i] = Calendar.getInstance().get(Calendar.YEAR) + offset + i;
			}
	    	  
	        dialog = new JDialog(f,"设置参数",true);
	        Container dialogPane = dialog.getContentPane();
	        
	        GridBagLayout gbl = new GridBagLayout();
	        dialogPane.setLayout(gbl);
	        
	        JLabel yearLabel = new JLabel("考试年份：");
	        dialogPane.add(yearLabel , new GBC(0,0,1,1).setAnchor(GridBagConstraints.CENTER).setFill(GridBagConstraints.BOTH)
	        		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
	        
	        JLabel tNumLabel = new JLabel("监考人数：");
	        dialogPane.add(tNumLabel , new GBC(0,1,1,1).setAnchor(GridBagConstraints.CENTER).setFill(GridBagConstraints.BOTH)
	        		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
	        
	        JLabel MBANumLabel = new JLabel("监考人数(MBA)：");
	        dialogPane.add(MBANumLabel , new GBC(3,1,1,1).setAnchor(GridBagConstraints.CENTER).setFill(GridBagConstraints.BOTH)
	        		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
	        
	        JLabel marksLabel = new JLabel("备注信息：");
	        dialogPane.add(marksLabel , new GBC(0,2,1,1).setAnchor(GridBagConstraints.CENTER).setFill(GridBagConstraints.BOTH)
	        		 		.setWeight(0, 0).setInsets(5).setIpad(0, 0));
	        
	        DefaultComboBoxModel yearMode=new DefaultComboBoxModel();
	        yearField = new JComboBox(yearMode);
	        dialogPane.add(yearField , new GBC(1,0,2,1).setAnchor(GridBagConstraints.CENTER).setFill(GridBagConstraints.BOTH)
    		 		.setWeight(1, 0).setInsets(5).setIpad(0, 0));
	        for (int i=0;i<yearNum.length ;i++){
	        	yearMode.addElement(yearNum[i]);
	        }
	        yearMode.setSelectedItem(year);
	        
	   	  	DefaultComboBoxModel tNumMode=new DefaultComboBoxModel();
	        tNumField = new JComboBox(tNumMode);
	        dialogPane.add(tNumField , new GBC(1,1,2,1).setAnchor(GridBagConstraints.CENTER).setFill(GridBagConstraints.BOTH)
    		 		.setWeight(1, 0).setInsets(5).setIpad(0, 0));
	        
	        DefaultComboBoxModel MBANumMode=new DefaultComboBoxModel();
	        MBANumField = new JComboBox(MBANumMode);
	        dialogPane.add(MBANumField , new GBC(4,1,2,1).setAnchor(GridBagConstraints.CENTER).setFill(GridBagConstraints.BOTH)
    		 		.setWeight(1, 0).setInsets(5).setIpad(0, 0));
	        for (int i=0;i<teacherNum.length ;i++){
	        	tNumMode.addElement(teacherNum[i]);
	        	MBANumMode.addElement(teacherNum[i]);
	        }
	        tNumMode.setSelectedItem(tNum);
	        MBANumMode.setSelectedItem(MBANum);
	        
	        marksField = new JTextField();
	        marksField.setText(marks);
	        dialogPane.add(marksField , new GBC(1,2,5,1).setAnchor(GridBagConstraints.CENTER).setFill(GridBagConstraints.BOTH)
    		 		.setWeight(1, 0).setInsets(5).setIpad(0, 0));
	        
	        JPanel buttonPanel = new JPanel();
	        buttonPanel.setLayout(new GridLayout(1,2));
	        JButton enter = new JButton("确定");
	        enter.addActionListener(this);
	        buttonPanel.add(enter);
	        JButton cancel = new JButton("取消");
	        cancel.addActionListener(this);
	        buttonPanel.add(cancel);
	        dialogPane.add(buttonPanel , new GBC(0,3,6,1).setAnchor(GridBagConstraints.CENTER).setFill(GridBagConstraints.BOTH)
    		 		.setWeight(1, 0).setInsets(5).setIpad(0, 0));
	        
	        dialog.setBounds(200,150,400,200);
	        dialog.show();
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
