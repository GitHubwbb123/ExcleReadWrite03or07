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
                content.setText(ExcelPOIUtil.readOneCell(filePath2003,Integer.valueOf(hang.getText().toString()),Integer.valueOf(lie.getText().toString())));
                break;
            case R.id.read2007:
                content.setText(ExcelPOIUtil.readOneCell(filePath2007,Integer.valueOf(hang.getText().toString()),Integer.valueOf(lie.getText().toString())));
                break;
            case R.id.write2003:
                ExcelPOIUtil.writeOneCell2003(filePath2003,Integer.valueOf(hang.getText().toString()),Integer.valueOf(lie.getText().toString()),content.getText().toString());
                break;
            case R.id.write2007:
                ExcelPOIUtil.writeOneCell2007(filePath2007,Integer.valueOf(hang.getText().toString()),Integer.valueOf(lie.getText().toString()),content.getText().toString());
            break;

        }

    }


}

