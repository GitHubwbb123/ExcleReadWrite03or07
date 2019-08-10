package com.mrwxb.exclereadwrite03or07;

import android.Manifest;
import android.os.Environment;
import android.support.annotation.NonNull;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import com.yanzhenjie.permission.AndPermission;
import com.yanzhenjie.permission.PermissionNo;
import com.yanzhenjie.permission.PermissionYes;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
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
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {
    File file;
    private Button read2003;
    private Button read2007;
    private Button write2003;
    private Button write2007;
    private EditText hang;
    private EditText lie;
    private EditText content;
    private String fileDierctory=Environment.getExternalStorageDirectory().getPath()+"/Test";
    private String filePath2003=Environment.getExternalStorageDirectory().getPath()+"/Test/my2003.xls";
    private String filePath2007=Environment.getExternalStorageDirectory().getPath()+"/Test/my2007.xlsx";
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        if(AndPermission.hasPermission(this,Manifest.permission.WRITE_EXTERNAL_STORAGE,Manifest.permission.READ_EXTERNAL_STORAGE)){

        }else{
            AndPermission.with(this)
                    .requestCode(100)
                    .permission(Manifest.permission.WRITE_EXTERNAL_STORAGE,Manifest.permission.READ_EXTERNAL_STORAGE)
                    .send();

        }
        file=new File(fileDierctory);
        if(!file.exists()){
            file.mkdir();
        }//先建立一个文件夹
        read2003=findViewById(R.id.read2003);
        read2007=findViewById(R.id.read2007);
        write2003=findViewById(R.id.write2003);
        write2007=findViewById(R.id.write2007);
        hang=findViewById(R.id.hang);
        lie=findViewById(R.id.lie);
        content=findViewById(R.id.content);
        read2003.setOnClickListener(this);
        read2007.setOnClickListener(this);
        write2003.setOnClickListener(this);
        write2007.setOnClickListener(this);

    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        AndPermission.onRequestPermissionsResult(this,requestCode,permissions,grantResults);
    }
    @PermissionYes(100)
    private void getPermission(List<String> grantedPermissions) {
        Toast.makeText(MainActivity.this, "接受权限", Toast.LENGTH_SHORT).show();
    }
    @PermissionNo(100)
    private void refusePermission(List<String> grantedPermissions) {
        Toast.makeText(MainActivity.this, "拒接了权限", Toast.LENGTH_SHORT).show();
    }

    @Override
    public void onClick(View v) {
        switch (v.getId()){
            case R.id.read2003:
                content.setText(readOneCell(filePath2003,Integer.valueOf(hang.getText().toString()),Integer.valueOf(lie.getText().toString())));
                break;
            case R.id.read2007:
                content.setText(readOneCell(filePath2007,Integer.valueOf(hang.getText().toString()),Integer.valueOf(lie.getText().toString())));
                break;
            case R.id.write2003:
                writeOneCell2003(filePath2003,Integer.valueOf(hang.getText().toString()),Integer.valueOf(lie.getText().toString()),content.getText().toString());
                break;
            case R.id.write2007:
                writeOneCell2007(filePath2007,Integer.valueOf(hang.getText().toString()),Integer.valueOf(lie.getText().toString()),content.getText().toString());
            break;

        }

    }

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
    public void writeOneCell2003(String filePath,int row,int cul,String content){
        File file=new File(filePath);
        if(!file.exists()) {
            try {
                InputStream myInput = getAssets().open("2003.xls");
                Workbook workbook = WorkbookFactory.create(myInput);
                final Sheet  mySheet = workbook.getSheetAt(0);
                Row myRow;
                for(int i=0;i<3000;i++){
                    mySheet.createRow(i);
                }//创建3000行够用了吧
                myRow=mySheet.getRow(row);
                Cell myCell = myRow.createCell(cul);
                myCell.setCellValue(content);
                OutputStream outputStream = new FileOutputStream(filePath);//这一步如果没有文件，会创建文件，但是不会创建文件夹，因此要先创建文件夹
                workbook.write(outputStream);
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
    public void writeOneCell2007(String filePath,int row,int cul,String content){
        File file=new File(filePath);
        if(!file.exists()) {
            try {
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

    /******
     * 读只有一个函数，不管2003，还是2007
     * @param filePath
     * @param row
     * @param cul
     * @return
     */
    private  String readOneCell(String filePath,int row,int cul) {
        String content = null;
        File file=new File(filePath);
        if(!file.exists()){
               Toast.makeText(MainActivity.this,"文件不存在",Toast.LENGTH_SHORT).show();
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

