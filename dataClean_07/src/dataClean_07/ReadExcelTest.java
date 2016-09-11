package dataClean_07;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;



public class ReadExcelTest {

    private static final String EXTENSION_XLS = "xls";
    private static final String EXTENSION_XLSX = "xlsx";

    /***
     * <pre>
     * 取得Workbook对象(xls和xlsx对象不同,不过都是Workbook的实现类)
     *   xls:HSSFWorkbook
     *   xlsx：XSSFWorkbook
     * @param filePath
     * @return
     * @throws IOException
     * </pre>
     */
    private static Workbook getWorkbook(String filePath) throws IOException {
        Workbook workbook = null;
        InputStream is = new FileInputStream(filePath);
        if (filePath.endsWith(EXTENSION_XLS)) {
            workbook = new HSSFWorkbook(is);
        } else if (filePath.endsWith(EXTENSION_XLSX)) {
            workbook = new XSSFWorkbook(is);
        }
        return workbook;
    }

    /**
     * 文件检查
     * @param filePath
     * @throws FileNotFoundException
     * @throws FileFormatException
     */
    private static void preReadCheck(String filePath) throws FileNotFoundException, FileFormatException {
        // 常规检查
        File file = new File(filePath);
        if (!file.exists()) {
            throw new FileNotFoundException("传入的文件不存在：" + filePath);
        }

        if (!(filePath.endsWith(EXTENSION_XLS) || filePath.endsWith(EXTENSION_XLSX))) {
            throw new FileFormatException("传入的文件不是excel");
        }
    }

    /**
     * 读取excel文件内容
     * @param filePath
     * @throws FileNotFoundException
     * @throws FileFormatException
     */
    public static void readExcel(String filePath) throws FileNotFoundException, FileFormatException {
        // 检查
        preReadCheck(filePath);
        // 获取workbook对象
        Workbook workbook = null;

        try {
            workbook = getWorkbook(filePath);
            // 读文件 一个sheet一个sheet地读取
            for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
                Sheet sheet = workbook.getSheetAt(numSheet);
                if (sheet == null) {
                    continue;
                }
                System.out.println("=======================" + sheet.getSheetName() + "=========================");

                int firstRowIndex = sheet.getFirstRowNum();
                int lastRowIndex = sheet.getLastRowNum();

                // 读取首行 即,表头
                Row firstRow = sheet.getRow(firstRowIndex);
                for (int i = firstRow.getFirstCellNum(); i <= firstRow.getLastCellNum(); i++) {
                    Cell cell = firstRow.getCell(i);
                    String cellValue = getCellValue(cell, true);
                    //System.out.print(" " + cellValue + "\t");
                }
                //System.out.println("");

                // 读取数据行
                for (int rowIndex = firstRowIndex + 1; rowIndex <= lastRowIndex; rowIndex++) {
                    Row currentRow = sheet.getRow(rowIndex);// 当前行
                    int firstColumnIndex = currentRow.getFirstCellNum(); // 首列
                    int lastColumnIndex = currentRow.getLastCellNum();// 最后一列
                    for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
                        Cell currentCell = currentRow.getCell(columnIndex);// 当前单元格
                        String currentCellValue = getCellValue(currentCell, true);// 当前单元格的值
                        //System.out.print(currentCellValue + "\t");
                    }
                    //System.out.println("");
                }
                System.out.println("======================================================");
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
               /* try {
                     workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }*/
            }
        }
    }

    /**
     * 取单元格的值
     * @param cell 单元格对象
     * @param treatAsStr 为true时，当做文本来取值 (取到的是文本，不会把“1”取成“1.0”)
     * @return
     */
    private static String getCellValue(Cell cell, boolean treatAsStr) {
        if (cell == null) {
            return "";
        }

        if (treatAsStr) {
            // 虽然excel中设置的都是文本，但是数字文本还被读错，如“1”取成“1.0”
            // 加上下面这句，临时把它当做文本来读取
            cell.setCellType(Cell.CELL_TYPE_STRING);
        }

        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else {
            return String.valueOf(cell.getStringCellValue());
        }
    }
    
    public static void main(String[] args) throws Exception{
    	DataManager testM = new DataManager();
        String[] attri = testM.getTestingClassifications();
        int indexNum=0;
        int indexNumSum=0;
        
        String[] label=new String[]{" ","小区名称","地区","街道","月租金","售价","房源均价","付押方式","税","租赁方式","房型","面积","楼层","朝向","装修程度","配置","房屋类型","物业费","物业公司","物业类","建筑年数","建筑年代","总户数","容积率","停车位","绿化率","开发商",
        		"周边配套","房产性质","出租数量","在售数量","浏览量","更新时间","门店","数据来源"};
        String [][] excel=new String[100000][label.length];
        int [] index=new int[label.length];
        //"小区名称" 0
        //,"地区" 1
        //"街道"2
        //"月租金"3
        //"售价"4
        //"房源均价"5
        //"付押方式"6
        //"税"7
        //"租赁方式"8
        //"房型"9
        //"面积"10
        //"楼层11
        //"朝向"12
        //"装修程度"13
        //"配置"14
        //"房屋类型"15
        //"物业费"16
        //"物业公司"17
        //"物业类"18"
        //建筑年数"19
        //"建筑年代"20
        //"总户数"21
        //"容积率"22
        //"停车位"23
        //"绿化率"24
        //"开发商"25
        //周边配套"26
        //房产性质",27
        //"出租数量"28
        //"在售数量"29
        //"浏览量",30
        //"更新时间"31
        //"门店"32
        //交通33
        //数据来源34
        for (int i1=0;i1<label.length;i1++){
        	excel[0][i1]=label[i1];
        }
        for (int m = 0; m <attri.length; m++) 
        {
        	String att = attri[m];
        	
        	String[] fPath=testM.getFilesPath(att);
        	for (int n = 0; n < fPath.length; n++) 
            {
        		String path=fPath[n];
        		//
                // 检查
                preReadCheck(path);
                // 获取workbook对象
                Workbook workbook = null;

                try {
                    workbook = getWorkbook(path);
                    // 读文件 一个sheet一个sheet地读取
                    for (int numSheet = 0; numSheet < workbook.getNumberOfSheets(); numSheet++) {
                        Sheet sheet = workbook.getSheetAt(numSheet);
                        if (sheet == null) {
                            continue;
                        }
                        System.out.println("=======================" + sheet.getSheetName() + "=========================");

                        int firstRowIndex = sheet.getFirstRowNum();
                        int lastRowIndex = sheet.getLastRowNum();

                        // 读取首行 即,表头
                        Row firstRow = sheet.getRow(firstRowIndex);
                        for (int i = firstRow.getFirstCellNum(); i <= firstRow.getLastCellNum(); i++) {
                            Cell cell = firstRow.getCell(i);
                    
                            String cellValue = getCellValue(cell, true);
                            if ((cellValue.equals("小区"))||((cellValue.contains("楼盘"))&&(!(cellValue.equals("楼盘简介"))))){
                            	index[i]=1;
                            } if ((cellValue.contains("地区"))||(cellValue.contains("商圈"))||(cellValue.contains("区域"))){
                            	index[i]=2;
                            } if(cellValue.contains("街道")){
                            	index[i]=3;
                            }if(((cellValue.contains("租金"))&&(!(cellValue.equals("日租金"))))||(cellValue.contains("价格"))){
                            	index[i]=4;
                            } if(cellValue.contains("售价")||(cellValue.contains("总价"))){
                            	index[i]=5;
                            } if(cellValue.contains("均价")){
                            	index[i]=6;
                            } if(cellValue.contains("付")){
                            	index[i]=7;
                            } if(cellValue.contains("税")){
                            	index[i]=8;
                            } if(cellValue.contains("租赁方")){
                            	index[i]=9;
                            } if((cellValue.contains("房型"))||(cellValue.contains("户型"))){
                            	index[i]=10;
                            } if(cellValue.contains("面积")){
                            	index[i]=11;
                            } if(cellValue.contains("楼层")){
                            	index[i]=12;
                            } if(cellValue.contains("朝向")){
                            	index[i]=13;
                            } if(cellValue.contains("装修")){
                            	index[i]=14;
                            } if((cellValue.contains("配置"))||(cellValue.contains("配套设施"))){
                            	index[i]=15;
                            } if(cellValue.contains("类型")){
                            	index[i]=16;
                            } if(cellValue.contains("物业费")){
                            	index[i]=17;
                            } if(cellValue.contains("物业公司")){
                            	index[i]=18;
                            } if(cellValue.contains("物业类")){
                            	index[i]=19;
                            } if((cellValue.contains("年数"))||(cellValue.contains("房龄"))){
                            	index[i]=20;
                            } if(cellValue.contains("年代")){
                            	index[i]=21;
                            } if(cellValue.contains("总户数")){
                            	index[i]=22;
                            } if(cellValue.contains("容积率")){
                            	index[i]=23;
                            } if(cellValue.contains("停车位")){
                            	index[i]=24;
                            } if(cellValue.contains("绿化率")){
                            	index[i]=25;
                            	
                            } if(cellValue.contains("开发商")){
                            	index[i]=26;
                            } if(cellValue.contains("周边配套")){
                            	index[i]=27;
                            } if(cellValue.contains("性质")){
                            	index[i]=28;
                            } if(cellValue.contains("出租数量")){
                            	index[i]=29;
                            } if(cellValue.contains("在售数量")){
                            	index[i]=30;
                            } if(cellValue.contains("浏览量")){
                            	index[i]=31;
                            } if((cellValue.contains("更新"))||(cellValue.contains("发布"))){
                            	index[i]=32;
                            }if(cellValue.contains("门店")){
                            	index[i]=33;
                            }if(cellValue.contains("交通")){
                            	index[i]=34;
                            }
                            System.out.print(" " + cellValue + "\t");
                        }
                        System.out.println("");

                        // 读取数据行
                        for (int rowIndex = firstRowIndex + 1; rowIndex <= lastRowIndex; rowIndex++) {
                            Row currentRow = sheet.getRow(rowIndex);// 当前行
                            int firstColumnIndex = currentRow.getFirstCellNum(); // 首列
                            int lastColumnIndex = currentRow.getLastCellNum();// 最后一列
                            for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++) {
                                Cell currentCell = currentRow.getCell(columnIndex);// 当前单元格
                                String currentCellValue = getCellValue(currentCell, true);// 当前单元格的值
                                excel[rowIndex+indexNumSum][index[columnIndex]]=currentCellValue;
                                //System.out.print(currentCellValue + "\t");
                                
                            }
                            excel[rowIndex+indexNumSum][excel[0].length-1]=att;

                           // System.out.println("");
                            indexNum++;
                        }
                        
                        System.out.println("======================================================");
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }/* finally {
                    if (workbook != null) {
                        try {
                             workbook.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }*/
        		//
                indexNumSum+=indexNum;
                System.out.println(indexNumSum);
                indexNum=0;
    }
        	
        }
        /*for (int q = 0; q < excel.length; q++) {
            for (int p= 0; p < excel[0].length; p++) {
                
            	System.out.print(excel[q][p] + " ");
            }
            System.out.print("\n");
        }*/
        
        File file = new File("/office_zhongjing/dataSet/processedData/New Text Document.txt");  //存放数组数据的文件
        
        FileWriter out = new FileWriter(file);  //文件写入流
       
        //将数组中的数据写入到文件中。每行各数据之间TAB间隔
        for(int q = 0; q < excel.length; q++){
         for(int p= 0; p < excel[0].length; p++){
        	 if(excel[q][p]==null){
           	  excel[q][p]="";
        	 }
          out.write(excel[q][p]+"?");
         }
         out.write("\n");
        }
        out.close();
    }

}