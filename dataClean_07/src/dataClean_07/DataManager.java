package dataClean_07;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
* 训练集管理器
*/
public class DataManager {

    private String[] testingFileClassifications;//训练语料分类集合
    private File testingTextDir;//训练语料存放目录
    private static String defaultPath = "/office_zhongjing/dataSet/dataCollect";
    
    
    public DataManager() 
    {
        testingTextDir = new File(defaultPath);
        if (!testingTextDir.isDirectory()) 
        {
            throw new IllegalArgumentException("测试语料库搜索失败！ [" +defaultPath + "]");
        }
        this.testingFileClassifications = testingTextDir.list();
    }
    
    /**
    * 返回训练文本类别，这个类别就是目录名
    * @return 训练文本类别
    */
    public String[] getTestingClassifications() 
    {
        return this.testingFileClassifications;
    }
    
    /**
    * 根据训练文本类别返回这个类别下的所有训练文本路径（full path）
    * @param classification 给定的分类
    * @return 给定分类下所有文件的路径（full path）
    */
    public String[] getFilesPath(String classification) 
    {
        File classDir = new File(testingTextDir.getPath() +File.separator +classification);
        String[] ret = classDir.list();
        for (int i = 0; i < ret.length; i++) 
        {
            ret[i] = testingTextDir.getPath() +File.separator +classification +File.separator +ret[i];
        }
        return ret;
    }
    
    /**
    * 返回给定路径的文本文件内容
    * @param filePath 给定的文本文件路径
    * @return 文本内容
    * @throws java.io.FileNotFoundException
    * @throws java.io.IOException
     * @throws BiffException 
    */
    public static String getText(String filePath) throws FileNotFoundException,IOException, BiffException 
    {     
	     //构建Workbook对象, 只读Workbook对象   
	     //直接从本地文件创建Workbook   
	      //从输入流创建Workbook   
	  
	        FileInputStream fis = new FileInputStream(filePath);     
	        StringBuilder sb = new StringBuilder();     
	        jxl.Workbook rwb = Workbook.getWorkbook(fis);     
	//一旦创建了Workbook，我们就可以通过它来访问   
	//Excel Sheet的数组集合(术语：工作表)，   
	//也可以调用getsheet方法获取指定的工资表   
	        Sheet[] sheet = rwb.getSheets();     
	        for (int i = 0; i < sheet.length; i++) {     
	            Sheet rs = rwb.getSheet(i);     
	            for (int j = 0; j < rs.getRows(); j++) {     
	               Cell[] cells = rs.getRow(j);     
	               for(int k=0;k<rs.getColumns();k++){
	               sb.append(cells[k].getContents());     
	            }     
	        }
	        }
	        fis.close();     
	        return sb.toString();     
	    }   
    
    /**
    * 返回训练文本集中所有的文本数目
    * @return 训练文本集中所有的文本数目
    */
    public int getTestingFileCount()
    {
        int ret = 0;
        for (int i = 0; i < testingFileClassifications.length; i++)
        {
            ret +=getTestingFileCountOfClassification(testingFileClassifications[i]);
        }
        return ret;
    }
    
    /**
    * 返回训练文本集中在给定分类下的训练文本数目
    * @param classification 给定的分类
    * @return 训练文本集中在给定分类下的训练文本数目
    */
    public int getTestingFileCountOfClassification(String classification)
    {
        File classDir = new File(testingTextDir.getPath() +File.separator +classification);
        return classDir.list().length;
    }
    
    /**
    * 返回给定分类中包含关键字／词的训练文本的数目
    * @param classification 给定的分类
    * @param key 给定的关键字／词
    * @return 给定分类中包含关键字／词的训练文本的数目
     * @throws BiffException 
    */
    public int getCountContainKeyOfClassification(String classification,String key) throws BiffException 
    {
        int ret = 0;
        try 
        {
            String[] filePath = getFilesPath(classification);
            for (int j = 0; j < filePath.length; j++) 
            {
                String text = getText(filePath[j]);
                if (text.contains(key)) 
                {
                    ret++;
                }
            }
        }
        catch (FileNotFoundException ex) 
        {
        		Logger.getLogger(DataManager.class.getName()).log(Level.SEVERE, null,ex);
        } 
        catch (IOException ex)
        {
            Logger.getLogger(DataManager.class.getName()).log(Level.SEVERE, null,ex);
        }
        return ret;
    }
}