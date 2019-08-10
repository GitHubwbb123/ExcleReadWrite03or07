package com.mrwxb.exclereadwrite03or07;

import android.content.res.AssetManager;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public  class ExcelPOIUtil {



    /**************
     现在用WorkbookFactory模式，不用管是2007还是2003了，因此注释掉
     * @param filePath
     * @param row
     * @param cul
     */
    /*private  String readOneCellExcel2003(String filePath,int row,int cul){
        try {
            InputStream inputStream=new FileInputStream(filePath);
            POIFSFileSystem poifsFileSystem=new POIFSFileSystem(inputStream);
            HSSFWorkbook hssfWorkbook=new HSSFWorkbook(poifsFileSystem);
            HSSFSheet hssfSheet=hssfWorkbook.getSheetAt(0);
            HSSFRow hssfRow=hssfSheet.getRow(row);
            HSSFCell hssfCell=hssfRow.getCell(cul);
            hssfCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            return hssfCell.getStringCellValue();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "";
    }

    */
/*
    private  String readOneCellExcel2007(String filePath,int row,int cul){
        try {
            InputStream inputStream=new FileInputStream(filePath);
            XSSFWorkbook xssfWorkbook=new XSSFWorkbook(inputStream);
            XSSFSheet xssfSheet=xssfWorkbook.getSheetAt(0);
            XSSFRow xssfRow=xssfSheet.getRow(row);
            XSSFCell xssfCell=xssfRow.getCell(cul);
            xssfCell.setCellType(XSSFCell.CELL_TYPE_STRING);
            return xssfCell.getStringCellValue();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "";
    }
    */
    /*
    private  void writeOneCellExcel2007(String filePath,int row,int cul){
        try {
            InputStream inputStream=new FileInputStream(filePath);
            XSSFWorkbook xssfWorkbook=new XSSFWorkbook(inputStream);
            XSSFSheet xssfSheet=xssfWorkbook.createSheet();
            XSSFRow xssfRow=xssfSheet.createRow(row);
            XSSFCell xssfCell=xssfRow.createCell(cul);
            xssfCell.setCellValue("我是刚建立的2007");
            OutputStream outputStream=new FileOutputStream(filePath);
            xssfWorkbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
*/
   /* private  void writeOneCellExcel2003(String filePath,int row,int cul){
        try {
            InputStream inputStream=new FileInputStream(filePath);
            POIFSFileSystem poifsFileSystem=new POIFSFileSystem(inputStream);
            HSSFWorkbook hssfWorkbook=new HSSFWorkbook(poifsFileSystem);
            HSSFSheet hssfSheet=hssfWorkbook.createSheet();
            HSSFRow hssfRow=hssfSheet.createRow(row);
            HSSFCell hssfCell=hssfRow.createCell(cul);
            hssfCell.setCellValue(content.getText(););
            OutputStream outputStream=new FileOutputStream(filePath);
            hssfWorkbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
*/
    /*
     * factory 模式
     * @param filePath
     * @param row
     * @param cul
     */
    public static void writeOneCell2003(String filePath,int row,int cul,String content){
        File file=new File(filePath);
        if(!file.exists()) {
            try {
                HSSFWorkbook hssfWorkbook=new HSSFWorkbook();
                final HSSFSheet hssfSheet=hssfWorkbook.createSheet("Sheet1");
                for(int i=0;i<3000;i++){
                    hssfSheet.createRow(i);
                }//创建3000行够用了吧
                HSSFRow hssfRow=hssfSheet.getRow(row);
                HSSFCell hssfCell=hssfRow.createCell(cul);
                hssfCell.setCellValue(content);
                OutputStream outputStream = new FileOutputStream(file);//这一步如果没有文件，会创建文件，但是不会创建文件夹，因此要先创建文件夹
                hssfWorkbook.write(outputStream);
                outputStream.flush();
                outputStream.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        else{
            try {
                InputStream myInput =new FileInputStream(filePath);
                Workbook workbook = WorkbookFactory.create(myInput);
                Sheet mySheet = workbook.getSheetAt(0);
                Row myRow=mySheet.getRow(row);
                Cell myCell = myRow.createCell(cul);
                myCell.setCellValue(content);
                OutputStream outputStream = new FileOutputStream(filePath);//这一步如果没有文件，会创建文件，但是不会创建文件夹，因此要先创建文件夹
                workbook.write(outputStream);
                myInput.close();
                outputStream.flush();
                outputStream.close();

            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return;
    }

    /*
     * factory 模式
     * @param filePath
     * @param row
     * @param cul
     */
    public static void writeOneCell2007(String filePath,int row,int cul,String content){
        File file=new File(filePath);
        if(!file.exists()) {
            try {
                XSSFWorkbook xssfWorkbook=new XSSFWorkbook();
                final XSSFSheet xssfSheet=xssfWorkbook.createSheet("Sheet1");
                for(int i=0;i<3000;i++){
                    xssfSheet.createRow(i);//只有创建的行，以后才可以写入，没有创建的行的单元是不能写入的，并且创建新的行，这一行所有单元会全部清空，因此在使用前用createSheet，以后就用getSheet，才不会把行数据全部清空
                }//创建3000行够用了吧
                XSSFRow xssfRow=xssfSheet.getRow(row);
                XSSFCell xssfCell=xssfRow.createCell(cul);//这里createCell/gerCell都可以，因为是写，不是读。
                xssfCell.setCellValue(content);
                OutputStream outputStream=new FileOutputStream(filePath);//这里用file或则filePath都可以，
                xssfWorkbook.write(outputStream);
                outputStream.flush();
                outputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
           /* try {
                InputStream myInput = getAssets().open("2007.xlsx");
                Workbook workbook = WorkbookFactory.create(myInput);
                final Sheet  mySheet = workbook.getSheetAt(0);
                Row myRow;
                        for(int i=0;i<3000;i++){
                            mySheet.createRow(i);
                        }//创建3000行够用了吧
                myRow=mySheet.createRow(row);
                Cell myCell = myRow.createCell(cul);
                myCell.setCellValue(content);
                OutputStream outputStream = new FileOutputStream(filePath);//这一步如果没有文件，会创建文件，但是不会创建文件夹，因此要先创建文件夹
                workbook.write(outputStream);
                outputStream.flush();
                outputStream.close();

            } catch (Exception e) {
                e.printStackTrace();
            }*/
            /***Workbook workbook = WorkbookFactory.create(myInput);得到的workbook，在新建时，建立的excel不能用，最后在workbook.write(outputStream)的时候写不进excel,文件是空的，
             * 只能用HSSFWorkbook hssfWorkbook=new HSSFWorkbook();得到的hssfWorkbook对象最后workbook.write(outputStream);新建文件，写出的输出流才能写进excel，只要新建成功，以后就随便用factory或则HSS,XSS，都可以读写
             * 备注：
             */
        }
        else{
            try {
                InputStream myInput =new FileInputStream(filePath);
                Workbook workbook = WorkbookFactory.create(myInput);
                Sheet mySheet = workbook.getSheetAt(0);
                Row myRow=mySheet.getRow(row);
                Cell myCell = myRow.createCell(cul);
                myCell.setCellValue(content);
                OutputStream outputStream = new FileOutputStream(filePath);//这一步如果没有文件，会创建文件，但是不会创建文件夹，因此要先创建文件夹
                workbook.write(outputStream);
                myInput.close();
                outputStream.flush();
                outputStream.close();

            } catch (Exception e) {
                e.printStackTrace();
            }



        }
        return;
    }

    /******
     * 读只有一个函数，不管2003，还是2007
     * @param filePath
     * @param row
     * @param cul
     * @return
     */
    public static String readOneCell(String filePath,int row,int cul) {
        String content = null;
        File file=new File(filePath);
        if(!file.exists()){
            //Toast.makeText(this,"文件不存在",Toast.LENGTH_SHORT).show();
        }
        else{
            try {
                InputStream myInput = new FileInputStream(filePath);
                Workbook workbook = WorkbookFactory.create(myInput);
                Sheet mySheet = workbook.getSheetAt(0);
                Row myRow = mySheet.getRow(row);
                Cell myCell = myRow.getCell(cul);
                myCell.setCellType(XSSFCell.CELL_TYPE_STRING);//先全部转换成String,因为POI分 cell.getStringCellValue()和get其他类型，如果格式不对就要报错推出
                content = myCell.getStringCellValue();
            } catch (Exception e) {
                e.printStackTrace();
            }

        }
        return content;//返回结果，没有返回空
    }


}
