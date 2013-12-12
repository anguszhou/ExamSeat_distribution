package com;

import java.util.ArrayList;    
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;     
import java.util.Map;
import java.util.Map.Entry;
import java.util.Random;
   
import com.jacob.activeX.ActiveXComponent;     
import com.jacob.com.ComThread;     
import com.jacob.com.Dispatch;     
import com.jacob.com.Variant;    
   
import org.apache.log4j.Logger;    
   
/**   
 *    
 *作用：利用jacob插件生成word 文件！    
 *    
 *
 */   
public class DOCWriter {    
   
    /** 日志记录器 */   
    static private Logger logger = Logger.getLogger(DOCWriter.class);    
   
    /** word文档   
     *    
     * 在本类中有两种方式可以进行文档的创建,<br>   
     * 第一种调用 createNewDocument   
     * 第二种调用 openDocument    
     */     
    private Dispatch document = null;     
   
    /** word运行程序对象 */     
    private ActiveXComponent word = null;     
   
    /** 所有word文档 */     
    private Dispatch documents = null;     
   
    /**   
     *  Selection 对象 代表窗口或窗格中的当前所选内容。 所选内容代表文档中选定（或突出显示）的区域，如果文档中没有选定任何内容，则代表插入点。   
     *  每个文档窗格只能有一个Selection 对象，并且在整个应用程序中只能有一个活动的 Selection 对象。   
     */   
    private Dispatch selection = null;     
   
    /**   
     *    
     * Range 对象 代表文档中的一个连续区域。 每个 Range 对象由一个起始字符位置和一个终止字符位置定义。   
     * 说明：与书签在文档中的使用方法类似，Range 对象在 Visual Basic 过程中用来标识文档的特定部分。   
     * 但与书签不同的是，Range对象只在定义该对象的过程运行时才存在。   
     * Range对象独立于所选内容。也就是说，您可以定义和处理一个范围而无需更改所选内容。还可以在文档中定义多个范围，但每个窗格中只能有一个所选内容。   
     */   
    private Dispatch range = null;    
   
    /**   
     * PageSetup 对象 该对象包含文档的所有页面设置属性（如左边距、下边距和纸张大小）。   
     */   
    private Dispatch pageSetup = null;    
   
    /** 文档中的所有表格对象 */   
    private Dispatch tables = null;    
   
    /** 一个表格对象 */   
    private Dispatch table = null;    
   
    /** 表格所有行对象 */   
    private Dispatch rows = null;    
   
    /** 表格所有列对象 */   
    private Dispatch cols = null;    
   
    /** 表格指定行对象 */   
    private Dispatch row = null;    
   
    /** 表格指定列对象 */   
    private Dispatch col = null;    
   
    /** 表格中指定的单元格 */   
    private Dispatch cell = null;    
        
    /** 字体 */   
    private Dispatch font = null;    
        
    /** 对齐方式 */   
    private Dispatch alignment = null;    
   
    /**   
     * 构造方法   
     */   
    public DOCWriter() {     
   
        if(this.word == null){    
            /* 初始化应用所要用到的对象实例 */   
            this.word = new ActiveXComponent("Word.Application");     
            /* 设置Word文档是否可见，true-可见false-不可见 */   
            this.word.setProperty("Visible",new Variant(false));    
            /* 禁用宏 */   
            this.word.setProperty("AutomationSecurity", new Variant(3));    
        }    
        if(this.documents == null){    
            this.documents = word.getProperty("Documents").toDispatch();    
        }    
        ComThread.InitSTA();
    }    
    
    public void setVisible(boolean flag){
    	this.word.setProperty("Visible",new Variant(flag));  
    }
   
    /**   
     * 设置页面方向和页边距   
     *   
     * @param orientation   
     *            可取值0或1，分别代表横向和纵向   
     * @param leftMargin   
     *            左边距的值   
     * @param rightMargin   
     *            右边距的值   
     * @param topMargin   
     *            上边距的值   
     * @param buttomMargin   
     *            下边距的值   
     */   
    public void setPageSetup(int orientation, int leftMargin,int rightMargin, int topMargin, int buttomMargin) {    
            
        logger.debug("设置页面方向和页边距...");    
        if(this.pageSetup == null){    
            this.getPageSetup();    
        }    
        Dispatch.put(pageSetup, "Orientation", orientation);    
        Dispatch.put(pageSetup, "LeftMargin", leftMargin);    
        Dispatch.put(pageSetup, "RightMargin", rightMargin);    
        Dispatch.put(pageSetup, "TopMargin", topMargin);    
        Dispatch.put(pageSetup, "BottomMargin", buttomMargin);    
    }    
   
    /**    
     * 打开文件    
     *    
     * @param inputDoc    
     *            要打开的文件，全路径    
     * @return Dispatch    
     *            打开的文件    
     */     
    public Dispatch openDocument(String inputDoc) {     
   
        logger.debug("打开Word文档...");    
        this.document = Dispatch.call(documents,"Open",inputDoc).toDispatch();    
        this.getSelection();    
        this.getRange();    
        this.getAlignment();    
        this.getFont();    
        this.getPageSetup();    
        return this.document;     
    }     
   
    /**   
     * 创建新的文件   
     *    
     * @return Dispache 返回新建文件   
     */   
    public Dispatch createNewDocument(){    
            
        logger.debug("创建新的文件...");    
        this.document = Dispatch.call(documents,"Add").toDispatch();    
        this.getSelection();    
        this.getRange();    
        this.getPageSetup();    
        this.getAlignment();    
        this.getFont();    
        return this.document;    
    }    
   
    /**    
     * 选定内容    
     * @return Dispatch 选定的范围或插入点    
     */     
    public Dispatch getSelection() {     
   
        logger.debug("获取选定范围的插入点...");    
        this.selection = word.getProperty("Selection").toDispatch();    
        return this.selection;     
    }     
   
    /**   
     * 获取当前Document内可以修改的部分<p><br>   
     * 前提条件：选定内容必须存在   
     *    
     * @param selectedContent 选定区域   
     * @return 可修改的对象   
     */   
    public Dispatch getRange() {    
   
        logger.debug("获取当前Document内可以修改的部分...");    
        this.range = Dispatch.get(this.selection, "Range").toDispatch();    
        return this.range;    
    }    
   
    /**   
     * 获得当前文档的文档页面属性   
     */   
    public Dispatch getPageSetup() {    
            
        logger.debug("获得当前文档的文档页面属性...");    
        if(this.document == null){    
            logger.warn("document对象为空...");    
            return this.pageSetup;    
        }    
        this.pageSetup = Dispatch.get(this.document, "PageSetup").toDispatch();    
        return this.pageSetup;    
    }    
   
    /**    
     * 把选定内容或插入点向上移动    
     * @param count 移动的距离    
     */     
    public void moveUp(int count) {     
   
        logger.debug("把选定内容或插入点向上移动...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection,"MoveUp");    
        }    
    }     
   
    /**    
     * 把选定内容或插入点向下移动    
     * @param count 移动的距离    
     */     
    public void moveDown(int count) {     
   
        logger.debug("把选定内容或插入点向下移动...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection,"MoveDown");    
        }    
    }     
   
    /**    
     * 把选定内容或插入点向左移动    
     * @param count 移动的距离    
     */     
    public void moveLeft(int count) {     
   
        logger.debug("把选定内容或插入点向左移动...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection,"MoveLeft");    
        }    
    }     
   
    /**    
     * 把选定内容或插入点向右移动    
     * @param count 移动的距离    
     */     
    public void moveRight(int count) {     
   
        logger.debug("把选定内容或插入点向右移动...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection,"MoveRight");    
        }    
    }    
        
    /**   
     * 回车键   
     */   
    public void enterDown(int count){    
            
        logger.debug("按回车键...");    
        for(int i = 0;i < count;i++) {    
            Dispatch.call(this.selection, "TypeParagraph");    
        }    
    }    
   
    /**    
     * 把插入点移动到文件首位置    
     */     
    public void moveStart() {     
   
        logger.debug("把插入点移动到文件首位置...");    
        Dispatch.call(this.selection,"HomeKey",new Variant(6));     
    }     
   
    /**    
     * 从选定内容或插入点开始查找文本    
     * @param selection 选定内容    
     * @param toFindText 要查找的文本    
     * @return boolean true-查找到并选中该文本，false-未查找到文本    
     */     
    public boolean find(String toFindText) {     
   
        logger.debug("从选定内容或插入点开始查找文本"+" 要查找内容：  "+toFindText);    
        /* 从selection所在位置开始查询 */   
        Dispatch find = Dispatch.call(this.selection,"Find").toDispatch();     
        /* 设置要查找的内容 */   
        Dispatch.put(find,"Text",toFindText);     
        /* 向前查找 */   
        Dispatch.put(find,"Forward","True");     
        /* 设置格式 */   
        Dispatch.put(find,"Format","True");     
        /* 大小写匹配 */   
        Dispatch.put(find,"MatchCase","True");     
        /* 全字匹配 */   
        Dispatch.put(find,"MatchWholeWord","True");     
        /* 查找并选中 */   
        return Dispatch.call(find,"Execute").getBoolean();     
    }     
   
    /**    
     * 把选定内容替换为设定文本    
     * @param selection 选定内容    
     * @param newText 替换为文本    
     */     
    public void replace(String newText) {     
   
        logger.debug("把选定内容替换为设定文本...");    
        /* 设置替换文本 */   
        Dispatch.put(this.selection,"Text",newText);     
    }     
   
    /**    
     * 全局替换    
     * @param selection 选定内容或起始插入点    
     * @param oldText 要替换的文本    
     * @param replaceObj 替换为文本   
     */     
    public void replaceAll(String oldText,Object replaceObj) {     
   
        logger.debug("全局替换...");    
        /* 移动到文件开头 */   
        moveStart();     
        /* 表格替换方式 */   
        String newText = (String) replaceObj;    
        /* 图片替换方式 */   
        if(oldText.indexOf("image") != -1 || newText.lastIndexOf(".bmp") != -1 || newText.lastIndexOf(".jpg") != -1 || newText.lastIndexOf(".gif") != -1){     
            while (find(oldText)) {     
                insertImage(newText);     
                Dispatch.call(this.selection,"MoveRight");     
            }     
            /* 正常替换方式 */   
        } else {    
            while (find(oldText)) {     
                replace(newText);     
                Dispatch.call(this.selection,"MoveRight");     
            }     
        }    
    }     
   
    /**    
     * 插入图片    
     * @param selection 图片的插入点    
     * @param imagePath 图片文件（全路径）    
     */     
    public void insertImage(String imagePath) {     
   
        logger.debug("插入图片...");    
        //Dispatch.call(this.selection, "TypeParagraph");    
        Dispatch pic = Dispatch.call(Dispatch.get(this.selection,"InLineShapes").toDispatch(),"AddPicture",imagePath).toDispatch();
        Dispatch.call(pic, "Select");
        Dispatch.put(pic, "Width", new Variant(60)); // 图片的宽度
        Dispatch.put(pic, "Height", new Variant(80)); // 图片的高度
        Dispatch.call(selection, "MoveRight"); 
    }     
   
    /**   
     * 合并表格   
     *   
     * @param selection 操作对象   
     * @param tableIndex 表格起始点   
     * @param fstCellRowIdx 开始行   
     * @param fstCellColIdx 开始列   
     * @param secCellRowIdx 结束行   
     * @param secCellColIdx 结束列   
     */   
    public void mergeCell(int tableIndex, int fstCellRowIdx, int fstCellColIdx, int secCellRowIdx, int secCellColIdx){    
   
        logger.debug("合并单元格...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return;    
        }    
        Dispatch fstCell = Dispatch.call(table, "Cell",new Variant(fstCellRowIdx), new Variant(fstCellColIdx)).toDispatch();    
        Dispatch secCell = Dispatch.call(table, "Cell",new Variant(secCellRowIdx), new Variant(secCellColIdx)).toDispatch();    
        Dispatch.call(fstCell, "Merge", secCell);    
    }    
   
    /**   
     * 想Table对象中插入数值
     *    
     * @param selection 插入点   
     * @param tableIndex 表格起始点   
     * @param list 数据内容   
     */  
    public void insertRowToTable(String[] strs , int row){    
         logger.debug("向Table对象中插入数据...");    
         if(strs == null || strs.length <= 0){    
             logger.warn("写出数据集为空...");    
             return;    
         }    
         if(this.table == null){    
             logger.warn("table对象为空...");    
             return;    
         }    
         
         for(int j = 0; j<strs.length; j++){    
             /* 遍历表格中每一个单元格，遍历次数与所要填入的内容数量相同 */   
        	  Dispatch cell = this.getCell(row+1, j+1);    
              /* 选中此单元格 */   
              Dispatch.call(cell, "Select");    
              /* 写出内容到此单元格中 */   
              Dispatch.put(this.selection, "Text", strs[j]); 
         }    
   
         //this.moveRight(1);    
    }
    /**   
     * 想Table对象中插入数值一组图片
     *    
     * @param selection 插入点   
     * @param tableIndex 表格起始点   
     * @param list 数据内容   
     */  
    public void insertToTable(String[] imgPaths , int row){    
         logger.debug("向Table对象中插入数据...");    
         if(imgPaths == null || imgPaths.length <= 0){    
             logger.warn("写出数据集为空...");    
             return;    
         }    
         if(this.table == null){    
             logger.warn("table对象为空...");    
             return;    
         }    
         
         for(int j = 0; j<imgPaths.length; j++){    
             /* 遍历表格中每一个单元格，遍历次数与所要填入的内容数量相同 */  
        	 if(imgPaths[j] == null || imgPaths[j].length() <= 0)
        		 continue;
             Dispatch cell = this.getCell(row+1, j+1);    
             /* 选中此单元格 */   
             Dispatch.call(cell, "Select"); 
             Dispatch.call(selection, "MoveLeft");
             
             Dispatch pic = Dispatch.call(Dispatch.get(this.selection,"InLineShapes").toDispatch(),"AddPicture",imgPaths[j]).toDispatch();
             Dispatch.call(pic, "Select");
             Dispatch.put(pic, "Width", new Variant(60)); // 图片的宽度
             Dispatch.put(pic, "Height", new Variant(80)); // 图片的高度
             //Dispatch.call(selection, "MoveRight");
             
              
         }    
   
         //this.moveRight(1);    
    }
    /**   
     * 想Table对象中插入数值<p>   
     *     参数形式：ArrayList<String[]>List.size()为表格的总行数<br>   
     *     String[]的length属性值应该与所创建的表格列数相同   
     *    
     * @param selection 插入点   
     * @param tableIndex 表格起始点   
     * @param list 数据内容   
     */   
    public void insertToTable(List<String[]> list){    
   
        logger.debug("向Table对象中插入数据...");    
        if(list == null || list.size() <= 0){    
            logger.warn("写出数据集为空...");    
            return;    
        }    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return;    
        }    
        for(int i = 0; i < list.size(); i++){    
            String[]  strs = list.get(i);    
            for(int j = 0; j<strs.length; j++){    
                /* 遍历表格中每一个单元格，遍历次数与所要填入的内容数量相同 */   
                Dispatch cell = this.getCell(i+1, j+1);    
                /* 选中此单元格 */   
                Dispatch.call(cell, "Select");    
                /* 写出内容到此单元格中 */   
                Dispatch.put(this.selection, "Text", strs[j]);    
                /* 移动游标到下一个位置 */   
            }    
            this.moveDown(1);    
        }    
        this.enterDown(1);    
    }    
   
    /**   
     * 在文档中正常插入文字内容   
     *    
     * @param selection 插入点   
     * @param list 数据内容   
     */   
    public void insertToDocument(List<String> list){    
   
        logger.debug("向Document对象中插入数据...");    
        if(list == null || list.size() <= 0){    
            logger.warn("写出数据集为空...");    
            return;    
        }    
        if(this.document == null){    
            logger.warn("document对象为空...");    
            return;    
        }    
        for(String str : list){    
            /* 写出至word中 */   
            //this.applyListTemplate(3, 2);    
            Dispatch.put(this.selection, "Text", str);    
            this.moveDown(1);    
            this.enterDown(1);    
        }    
    }    
    /**   
     * 在文档中正常插入文字内容   
     *    
     * @param selection 插入点   
     * @param list 数据内容   
     */   
    public void insertToDocument(String str){    
   
        logger.debug("向Document对象中插入数据...");    
        if(str == null ){    
            logger.warn("写出数据集为空...");    
            return;    
        }    
        if(this.document == null){    
            logger.warn("document对象为空...");    
            return;    
        }    
        Dispatch.put(this.selection, "Text", str);    
        this.moveDown(1);    
        this.enterDown(1);    
          
    }    
   
    /**   
     * 创建新的表格   
     *    
     * @param selection 插入点   
     * @param document 文档对象   
     * @param rowCount 行数   
     * @param colCount 列数   
     * @param width 边框数值 0浅色1深色   
     * @return 新创建的表格对象   
     */   
    public Dispatch createNewTable(int rowCount, int colCount, int width){    
   
        logger.debug("创建新的表格...");    
        if(this.tables == null){    
            this.getTables();    
        }    
        this.getRange();    
        if(rowCount > 0 && colCount > 0){    
            this.table = Dispatch.call(this.tables,"Add",this.range,new Variant(rowCount),new Variant(colCount),new Variant(width)).toDispatch();    
        }    
        /* 返回新创建表格 */   
        return this.table;    
    }    
   
    /**   
     * 获取Document对象中的所有Table对象   
     *    
     * @return 所有Table对象   
     */   
    public Dispatch getTables(){    
   
   
        logger.debug("获取所有表格对象...");    
        if(this.document == null){    
            logger.warn("document对象为空...");    
            return this.tables;    
        }    
        this.tables = Dispatch.get(this.document, "Tables").toDispatch();    
        return this.tables;    
    }    
        
    /**   
     * 获取Document中Table的数量   
     *    
     * @return 表格数量   
     */   
    public int getTablesCount(){    
            
        logger.debug("获取文档中表格数量...");    
        if(this.tables == null){    
            this.getTables();    
        }    
        return Dispatch.get(tables, "Count").getInt();    
            
    }    
   
    /**   
     * 获取指定序号的Table对象   
     *    
     * @param tableIndex Table序列   
     * @return   
     */   
    public Dispatch getTable(int tableIndex){    
   
        logger.debug("获取指定表格对象...");    
        if(this.tables == null){    
            this.getTables();    
        }    
        if(tableIndex >= 0){    
            this.table = Dispatch.call(this.tables, "Item", new Variant(tableIndex)).toDispatch();    
        }    
        return this.table;    
    }    
   
    /**   
     * 获取表格的总列数   
     *    
     * @return 总列数   
     */   
    public int getTableColumnsCount() {    
   
        logger.debug("获取表格总行数...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return 0;    
        }    
        return Dispatch.get(this.cols,"Count").getInt();    
    }    
   
    /**   
     * 获取表格的总行数   
     *    
     * @return 总行数   
     */   
    public int getTableRowsCount(){    
   
        logger.debug("获取表格总行数...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return 0;    
        }    
        return Dispatch.get(this.rows,"Count").getInt();    
    }    
    /**   
     * 获取表格列对象   
     *    
     * @return 列对象   
     */   
    public Dispatch getTableColumns() {    
   
        logger.debug("获取表格行对象...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return this.cols;    
        }    
        this.cols = Dispatch.get(this.table,"Columns").toDispatch();    
        return this.cols;    
    }    
   
   
    /**   
     * 获取表格的行对象   
     *    
     * @return 总行数   
     */   
    public Dispatch getTableRows(){    
   
        logger.debug("获取表格总行数...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return this.rows;    
        }    
        this.rows = Dispatch.get(this.table,"Rows").toDispatch();    
        return this.rows;    
    }    
   
    /**   
     * 获取指定表格列对象   
     *    
     * @return 列对象   
     */   
    public Dispatch getTableColumn(int columnIndex) {    
   
        logger.debug("获取指定表格行对象...");    
        if(this.cols == null){    
            this.getTableColumns();    
        }    
        if(columnIndex >= 0){    
            this.col = Dispatch.call(this.cols, "Item", new Variant(columnIndex)).toDispatch();    
        }    
        return this.col;    
    }    
   
    /**   
     * 获取表格中指定的行对象   
     *    
     * @param rowIndex 行序号   
     * @return 行对象   
     */   
    public Dispatch getTableRow(int rowIndex){    
   
        logger.debug("获取指定表格总行数...");    
        if(this.rows == null){    
            this.getTableRows();    
        }    
        if(rowIndex >= 0){    
            this.row = Dispatch.call(this.rows, "Item", new Variant(rowIndex)).toDispatch();    
        }    
        return this.row;    
    }    
        
    /**   
     * 自动调整表格   
     */   
    public void autoFitTable() {    
            
        logger.debug("自动调整表格...");    
        int count = this.getTablesCount();    
        for (int i = 0; i < count; i++) {    
            Dispatch table = Dispatch.call(tables, "Item", new Variant(i + 1)).toDispatch();    
            Dispatch cols = Dispatch.get(table, "Columns").toDispatch();    
            Dispatch.call(cols, "AutoFit");    
        }    
    }    
   
    /**   
     * 获取当前文档中，表格中的指定单元格   
     *   
     * @param CellRowIdx  单元格所在行   
     * @param CellColIdx 单元格所在列   
     * @return 指定单元格对象   
     */   
    public Dispatch getCell(int cellRowIdx, int cellColIdx) {    
   
            
        logger.debug("获取当前文档中，表格中的指定单元格...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return this.cell;    
        }    
        if(cellRowIdx >= 0 && cellColIdx >=0){    
            this.cell = Dispatch.call(this.table, "Cell", new Variant(cellRowIdx),new Variant(cellColIdx)).toDispatch();    
        }    
        return this.cell;    
    }    
   
    /**   
     * 设置文档标题   
     *    
     * @param title 标题内容   
     */   
    public void setTitle(String title){    
            
        logger.debug("设置文档标题...");    
        if(title == null || "".equals(title)){    
            logger.warn("文档标题为空...");    
            return;    
        }    
        Dispatch.call(this.selection, "TypeText", title);     
    }    
        
    /**   
     * 设置当前表格线的粗细   
     *   
     * @param width   
     *        width范围：1<w<13,如果是0，就代表没有框   
     */   
    public void setTableBorderWidth(int width) {    
   
        logger.debug("设置当前表格线的粗细...");    
        if(this.table == null){    
            logger.warn("table对象为空...");    
            return;    
        }    
        /*   
         * 设置表格线的粗细 1：代表最上边一条线 2：代表最左边一条线 3：最下边一条线 4：最右边一条线 5：除最上边最下边之外的所有横线   
         * 6：除最左边最右边之外的所有竖线 7：从左上角到右下角的斜线 8：从左下角到右上角的斜线   
         */   
        Dispatch borders = Dispatch.get(table, "Borders").toDispatch();    
        Dispatch border = null;    
        for (int i = 1; i < 7; i++) {    
            border = Dispatch.call(borders, "Item", new Variant(i)).toDispatch();    
            if (width != 0) {    
                Dispatch.put(border, "LineWidth", new Variant(width));    
                Dispatch.put(border, "Visible", new Variant(true));    
            } else if (width == 0) {    
                Dispatch.put(border, "Visible", new Variant(false));    
            }    
        }    
    }    
        
    /**   
     * 对当前selection设置项目符号和编号   
     * @param tabIndex   
     *     1: 项目编号   
     *     2: 编号   
     *     3: 多级编号   
     *     4: 列表样式   
     * @param index   
     *     0:表示没有 ,其它数字代表的是该Tab页中的第几项内容   
     */   
    public void applyListTemplate(int tabIndex,int index){    
   
        logger.debug("对当前selection设置项目符号和编号...");    
        /* 取得ListGalleries对象列表 */   
        Dispatch listGalleries = Dispatch.get(this.word, "ListGalleries").toDispatch();    
        /* 取得列表中一个对象 */   
        Dispatch listGallery = Dispatch.call(listGalleries, "Item", new Variant(tabIndex)).toDispatch();    
        Dispatch listTemplates = Dispatch.get(listGallery, "ListTemplates").toDispatch();    
        if(this.range == null){    
            this.getRange();    
        }    
        Dispatch listFormat = Dispatch.get(this.range, "ListFormat").toDispatch();    
        Dispatch.call(listFormat,"ApplyListTemplate",Dispatch.call(listTemplates, "Item", new Variant(index)), new Variant(true),new Variant(1),new Variant(0));    
    }    
        
    /**   
     * 增加文档目录   
     *   
     * 目前采用固定参数方式，以后可以动态进行调整   
     */   
    public void addTablesOfContents()    
    {    
      /* 取得ActiveDocument、TablesOfContents、range对象 */   
      Dispatch ActiveDocument = word.getProperty("ActiveDocument").toDispatch();    
      Dispatch TablesOfContents = Dispatch.get(ActiveDocument,"TablesOfContents").toDispatch();    
      Dispatch range = Dispatch.get(this.selection, "Range").toDispatch();    
      /* 增加目录 */     
      Dispatch.call(TablesOfContents,"Add",range,new Variant(true),new Variant(1),new Variant(3),new Variant(true),new Variant(""),new Variant(true),new Variant(true));    
        
    }    
   
        
    /**   
     * 设置当前Selection 位置方式   
     * @param selectedContent 0－居左；1－居中；2－居右。   
     */   
    public void setAlignment(int alignmentType) {    
            
        logger.debug("设置当前Selection 位置方式...");    
        if(this.alignment == null){    
            this.getAlignment();    
        }    
        Dispatch.put(this.alignment, "Alignment", alignmentType);    
    }    
        
    /**   
     * 获取当前选择区域的对齐方式   
     *    
     * @return 对其方式对象   
     */   
    public Dispatch getAlignment(){    
            
        logger.debug("获取当前选择区域的对齐方式...");    
        if(this.selection == null){    
            this.getSelection();    
        }    
        this.alignment = Dispatch.get(this.selection, "ParagraphFormat").toDispatch();    
        return this.alignment;    
    }    
        
    /**   
     * 获取字体对象   
     *    
     * @return 字体对象   
     */   
    public Dispatch getFont(){    
            
        logger.debug("获取字体对象...");    
        if(this.selection == null){    
            this.getSelection();    
        }    
        this.font = Dispatch.get(this.selection, "Font").toDispatch();    
        return this.font;    
    }    
        
    /**   
     * 设置选定内容的字体 注：在调用此方法前，选定区域对象selection必须存在   
     *   
     * @param fontName   
     *            字体名称，例如 "宋体"   
     * @param isBold   
     *            粗体   
     * @param isItalic   
     *            斜体   
     * @param isUnderline   
     *            下划线   
     * @param rgbColor   
     *            颜色，例如"255,255,255"   
     * @param fontSize   
     *            字体大小   
     * @param Scale   
     *            字符间距，百分比值。例如 70代表缩放为70%   
     */   
    public void setFontScale(String fontName, boolean isBold, boolean isItalic, boolean isUnderline, String rgbColor, int Scale, int fontSize) {    
            
        logger.debug("设置字体...");    
        Dispatch.put(this.font, "Name", fontName);    
        Dispatch.put(this.font, "Bold", isBold);    
        Dispatch.put(this.font, "Italic", isItalic);    
        Dispatch.put(this.font, "Underline", isUnderline);    
        Dispatch.put(this.font, "Color", rgbColor);    
        Dispatch.put(this.font, "Scaling", Scale);    
        Dispatch.put(this.font, "Size", fontSize);    
    }    
    
    public void save() {          
		 Dispatch.call(document, "Save");      
	 } 
    
    /**    
     * 保存文件    
     * @param outputPath 输出文件（包含路径）    
     */     
    public void saveAs(String outputPath) {     
   
        logger.debug("保存文件...");    
        if(this.document == null){    
            logger.warn("document对象为空...");    
            return;    
        }    
        if(outputPath ==null || "".equals(outputPath)){    
            logger.warn("文件保存路径为空...");    
            return;    
        }    
        Dispatch.call(this.document,"SaveAs",outputPath);     
    }    
        
    public void saveAsHtml(String htmlFile){    
        Dispatch.invoke(this.document,"SaveAs",Dispatch.Method, new Object[]{htmlFile,new Variant(8)}, new int[1]);    
    }    
   
    /**    
     * 关闭文件    
     * @param document 要关闭的文件    
     */     
    public void close() {    
   
        logger.debug("关闭文件...");    
        if(document == null){    
            logger.warn("document对象为空...");    
            return;    
        }    
        Dispatch.call(document,"Close",new Variant(0));     
    }    
   
    /**   
     * 列印word文件   
     *   
     */   
    public void printFile(){    
        logger.debug("打印文件...");    
        if(document == null){    
            logger.warn("document对象为空...");    
            return;    
        }    
        Dispatch.call(document,"PrintOut");    
    }    
   
    /**    
     * 退出程序    
     */     
    public void quit() {     
   
        logger.debug("退出程序");    
        word.invoke("Quit",new Variant[0]);     
        ComThread.Release();     
    }    
   
    /**
     * 换页
     * @throws Exception
     */
    public void nextPage()
        throws Exception
    {
      Dispatch.call(this.selection, "InsertBreak");
    }
    
    public static void main(String args[]) throws Exception{    
        DOCWriter writer = new DOCWriter(); 
        
        writer.createNewDocument(); 
        writer.saveAs("E://TestDocBenjamin.doc"); 
        writer.setPageSetup(1, 5, 5, 5, 5);
        writer.setAlignment(1);
        writer.setVisible(true); 
        
        String str = "陕西省硕士生考试统考试题使用申报表";
        writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 20);  
        writer.insertToDocument(str);
        
        writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 15);  
        str = "考点：考试代码  考点名称(盖章)           负责人:              统计人:  ";
        writer.insertToDocument(str);
        
        Map<Integer , Integer[]> calcu = new LinkedHashMap<Integer , Integer[]>();
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
        
        Random rd = new Random();
        int num = 30;
        List<String[]> listTab = new ArrayList<String[]>();
        String[] list = new String[num+1];
        list[0] = "科目\\需用袋数\\份数";
        for (int i = 1; i < num+1; i++) {
			list[i] = String.valueOf(i);
		}
        listTab.add(list);
        for(Entry<Integer, String> code : examCode.entrySet()){
        	list = new String[num+1];
        	list[0] = code.getValue()+"("+code.getKey()+")";
        	 for (int i = 1; i < num+1; i++) {
     			list[i] = String.valueOf(rd.nextInt(60));
     		}
        	 listTab.add(list);
        }
        writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 10); 
        writer.createNewTable(examCode.size()+1, num+8, 1);
        for (int i = 0; i < examCode.size()+1; i++) {
			writer.mergeCell(0, i+1, 1, i+1, 8);
		}
        //writer.mergeCell(0, 1, 1, 1, 3);
        for (int i = 0; i < examCode.size()+1; i++) {
        	writer.insertRowToTable(listTab.get(i), i);
		}
        //writer.insertToTable(listTab);
        
        /*
        List<String> list = new ArrayList<String>(); 
        String str = "照片对照表";
        writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 20);  
        writer.insertToDocument(str);
        
        writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 10);  
        str = "第1考场        第一单元     地点：沙发上了飞机阿拉斯加看到发牢骚";
        writer.insertToDocument(str);
        
        	 writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 8); 
             writer.createNewTable(15, 10, 0);
             
             for(int m = 0 ; m<3 ; m++){
            	 String[] imgs = new String[10];
                 for (int i = 0; i < imgs.length; i++) {
         			imgs[i] = "E:\\"+i+".jpg";
         			//writer.insertImage("E:\\a.jpg");
         			
         		}
                 writer.insertToTable(imgs, m*5);
                 
                 String str1[] = new String[10];  
                 for (int i = 0; i < 10; i++) {
    				str1[i] = "123456789";
    			 }
                 writer.insertRowToTable(str1, m*5+1);   
                 for (int i = 0; i < 10; i++) {
     				str1[i] = "撒旦发生的放地方";
     			 }
                 writer.insertRowToTable(str1, m*5+2);
                  for (int i = 0; i < 10; i++) {
      				str1[i] = "4566";
      			 }
                  writer.insertRowToTable(str1, m*5+3);
                   for (int i = 0; i < 10; i++) {
       				str1[i] = "撒旦法斯蒂芬斯蒂芬斯蒂芬";
       			 }
                   writer.insertRowToTable(str1, m*5+4);
             }
             writer.moveDown(1);
             writer.enterDown(1);
             writer.nextPage();
             
             String str2 = "照片对照表";
             writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 20);  
             writer.insertToDocument(str2);
             
             writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 10);  
             str2 = "第1考场        第一单元     地点：沙发上了飞机阿拉斯加看到发牢骚";
             writer.insertToDocument(str2);
             
             	 writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 8); 
                  writer.createNewTable(15, 10, 0);
                  
                  for(int m = 0 ; m<3 ; m++){
                 	 String[] imgs = new String[10];
                      for (int i = 0; i < imgs.length; i++) {
              			imgs[i] = "E:\\"+i+".jpg";
              			//writer.insertImage("E:\\a.jpg");
              			
              		}
                      writer.insertToTable(imgs, m*5);
                      
                      String str1[] = new String[10];  
                      for (int i = 0; i < 10; i++) {
         				str1[i] = "123456789";
         			 }
                      writer.insertRowToTable(str1, m*5+1);   
                      for (int i = 0; i < 10; i++) {
          				str1[i] = "撒旦发生的放地方";
          			 }
                      writer.insertRowToTable(str1, m*5+2);
                       for (int i = 0; i < 10; i++) {
           				str1[i] = "4566";
           			 }
                       writer.insertRowToTable(str1, m*5+3);
                        for (int i = 0; i < 10; i++) {
            				str1[i] = "撒旦法斯蒂芬斯蒂芬斯蒂芬";
            			 }
                        writer.insertRowToTable(str1, m*5+4);
                  }
                  */
            
        /*
        List<String[]> listTable = new ArrayList<String[]>();    
        for(int i = 0 ; i<3; i++){    
            String str1[] = new String[8];    
            str1[0] = "106984611101917"; 
            str1[1] = "士大夫" ;
            str1[2] = "106984611101917" ; 
            str1[3] = "士大夫" ;
            str1[4] = "106984611101917" ; 
            str1[5] = "士大夫" ;
            str1[6] = "106984611101917" ; 
            str1[7] = "士大夫" ;
            listTable.add(str1);    
        }    
        
        String str2[] = new String[8];    
        str2[0] = "106984611101917" ; 
        str2[1] = "士大" ;
        str2[2] = "106984611101917" ; 
        str2[3] = "士大撒旦发生的" ;
        str2[4] = "106984611101917" ; 
        str2[5] = "士大夫" ;
        str2[6] = "106984611101917" ; 
        str2[7] = "士大夫斯蒂芬斯蒂" ;
        listTable.add(str2);
        
        String str1[] = new String[8];    
        str1[0] = "106984611101917" ; 
        str1[1] = "士大夫受到法" ;
        str1[2] = "106984611101917" ; 
        str1[3] = "士大夫" ;
        str1[4] ="106984611101917"; 
        str1[5] = "士大夫" ;
        str1[6] = "106984611101917"; 
        str1[7] = "士大夫" ;
        listTable.add(str1);
        
        writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 18);  
        writer.createNewTable(5, 8, 0);    
        writer.insertToTable(listTable); 
        
        
        writer.nextPage();
        writer.setFontScale("宋体", false, false,false, "0,0,0,0", 100, 20);  
        str = "第2考场           人数：30     地点：沙发上了飞机阿拉斯加看到发牢骚";
        writer.insertToDocument(str);
        */
        writer.save();
        
//      writer.close();    
    }    
   
}  
