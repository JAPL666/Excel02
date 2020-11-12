package com.warma;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.text.DecimalFormat;
import java.util.Objects;
import java.util.Stack;

public class Utils {

    public static void findFiles(String path){

        Stack<String> stack=new Stack<>();
        getFilePath(stack,path);

        while (!stack.empty()) {
            String str= stack.pop();
            if(!str.contains("System Volume Information")){
                getFilePath(stack, str);
            }
        }

    }
    private static void getFilePath(Stack<String> stack,String pathName){
        //获取文件列表
        File[] file=new File(pathName).listFiles();
        if(file!=null){
            for (int i = 0; i< Objects.requireNonNull(file).length; i++){
                String path=file[i].getPath();
                //判断是否文件
                if(new File(path).isFile()){
                    boolean a=path.toLowerCase().endsWith(".xlsx");
                    boolean b=path.toLowerCase().endsWith(".xls");
                    if(a||b){
                        //System.out.println(path);
                        readExcel(path);
                    }
                }else{
                    //文件夹进栈
                    stack.push(path);
                }
            }
        }
    }
    public static void readExcel(String path){
        try {
            File file = new File(path);

            FileInputStream in = new FileInputStream(file);
            XSSFWorkbook wb = new XSSFWorkbook(in);

            Sheet sheet = wb.getSheetAt(0); //取得第一个表单
            int firstRowNum = sheet.getFirstRowNum()+3;//获取第一个数字
            int lastRowNum = sheet.getLastRowNum()-3;//获取最后一个数字

            for (int i = firstRowNum; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);//获取行

                String cell_3;
                Cell cell_ = row.getCell(3);

                //判断是否是数字
                if(cell_.getCellType().toString().equals("NUMERIC")){
                    //科学计数法转字符串
                    DecimalFormat df = new DecimalFormat("0");
                    cell_3 = df.format(cell_.getNumericCellValue());
                }else{
                    cell_3=cell_.toString().trim();
                }
                System.out.print(cell_3+"\n");

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
