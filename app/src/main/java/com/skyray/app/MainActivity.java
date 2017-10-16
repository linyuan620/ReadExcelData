package com.skyray.app;

import java.io.FileInputStream;
import java.io.InputStream;

import android.os.Bundle;
import android.app.Activity;
import android.text.method.ScrollingMovementMethod;
import android.widget.TextView;

import jxl.*;

public class MainActivity extends Activity {
    TextView txt = null;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        txt = (TextView)findViewById(R.id.txt_show);
        txt.setMovementMethod(ScrollingMovementMethod.getInstance());
        readExcel();
    }

    public void readExcel() {
        try {
            InputStream is = new FileInputStream("/motion.xls");

            Workbook book = Workbook.getWorkbook(is);

            int num = book.getNumberOfSheets();
            txt.setText("the num of sheets is " + num+ "\n");

            Sheet sheet = book.getSheet(0);
            int Rows = sheet.getRows();
            int Cols = sheet.getColumns();
            txt.append("the name of sheet is " + sheet.getName() + "\n");
            txt.append("total rows is " + Rows + "\n");
            txt.append("total cols is " + Cols + "\n");
            for (int i = 0; i < Cols; ++i) {
                for (int j = 0; j < Rows; ++j) {

                    txt.append("contents:" + sheet.getCell(i,j).getContents() + "\n");
                }
            }
            book.close();
        } catch (Exception e) {
            System.out.println(e);
        }
    }

}