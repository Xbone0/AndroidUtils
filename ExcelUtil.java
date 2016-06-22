package org.hahaqu.app.lockbankaccount.util;

import android.util.Log;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import jxl.Cell;
import jxl.CellView;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * Created by crossbones on 16/6/6.
 */
public class ExcelUtil {
    private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    public InputStream inputStream = null;
    private File inputFile = null;
    private File outputFile = null;
    private static ExcelUtil instance = null;

    public static synchronized ExcelUtil getInstance() {
        if (instance == null) {
            instance = new ExcelUtil();
        }
        return instance;
    }

    public void setInputFile(File inputFile) {
        this.inputFile = inputFile;
    }

    public void setOutputFile(File outputFile) {
        this.outputFile = outputFile;
    }

    public void setInputStream(InputStream inputFile) {
        this.inputStream = inputFile;
    }

    public ArrayList<String[]> ReadExcelByColumn() {
        Workbook workbook = null;
        try {
            if (inputFile == null) {
                workbook = Workbook.getWorkbook(inputStream);
            } else {
                workbook = Workbook.getWorkbook(inputFile);
            }
        } catch (Exception e) {
            e.printStackTrace();
            Log.e("EXCELUTIL", "Exception" + e.getMessage());
        }
        ArrayList contents = new ArrayList<String[]>();
        String[] content = null;
        // Get the first sheet
        Sheet sheet = workbook.getSheet(0);
        int columns = sheet.getColumns();
        int rows = sheet.getRows();
        // Loop over  column and lines
        for (int j = 0; j < columns; j++) {
            content = new String[rows];
            for (int i = 0, k = 0; i < rows; i++, k++) {
                Cell cell = sheet.getCell(j, i);
                String string = cell.getContents();
                content[k] = string;
            }
            contents.add(content);
        }
        return contents;
    }

    public ArrayList<String[]> ReadExcelByColumn(int column_start, int column_end, int row_start, int row_end) {

        Workbook workbook = null;
        try {
            if (inputFile == null) {
                workbook = Workbook.getWorkbook(inputStream);
            } else {
                workbook = Workbook.getWorkbook(inputFile);
            }
        } catch (Exception e) {
            e.printStackTrace();
            Log.e("EXCELUTIL", "Exception" + e.getMessage());
        }
        ArrayList contents = new ArrayList<String[]>();
        String[] content = null;
        System.out.println("columns:" + column_end + "rows:" + row_end);
        // Get the first sheet
        Sheet sheet = workbook.getSheet(0);
        int columns = column_end - column_start + 1;
        int rows = row_end - row_start + 1;
        if (rows > sheet.getRows() || columns > sheet.getColumns()) {
            System.out.println("columns or rows out of index");
            return null;
        }
        // Loop over  column and lines
        for (int j = column_start; j <= column_end; j++) {
            content = new String[rows];
            for (int i = row_start, k = 0; i <= row_end; i++, k++) {
                Cell cell = sheet.getCell(j, i);
                String string = cell.getContents();
                content[k] = string;
            }
            contents.add(content);
        }
        return contents;
    }

    public ArrayList<String[]> ReadExcelByRow(int column_start, int column_end, int row_start, int row_end) {

        Workbook workbook = null;
        try {
            if (inputFile == null) {
                workbook = Workbook.getWorkbook(inputStream);
            } else {
                workbook = Workbook.getWorkbook(inputFile);
            }
        } catch (Exception e) {
            e.printStackTrace();
            Log.e("EXCELUTIL", "Exception" + e.getMessage());
        }
        ArrayList contents = new ArrayList<String[]>();
        String[] content = null;
        System.out.println("columns:" + column_end + "rows:" + row_end);
        // Get the first sheet
        Sheet sheet = workbook.getSheet(0);
        int columns = column_end - column_start + 1;
        int rows = row_end - row_start + 1;
        if (rows > sheet.getRows() || columns > sheet.getColumns()) {
            System.out.println("columns or rows out of index");
            return null;
        }
        // Loop over  column and lines
        for (int i = row_start; i <= row_end; i++) {
            content = new String[rows];
            for (int j = column_start, k = 0; j <= column_end; j++, k++) {
                Cell cell = sheet.getCell(j, i);
                String string = cell.getContents();
                content[k] = string;
            }
            contents.add(content);
        }
        return contents;
    }

    /**
     * 返回行列数
     *
     * @return
     */
    public int[] getExcelSize() {
        Workbook workbook = null;
        try {
            if (inputFile == null) {
                workbook = Workbook.getWorkbook(inputStream);
            } else {
                workbook = Workbook.getWorkbook(inputFile);
            }
        } catch (Exception e) {
            e.printStackTrace();
            Log.e("EXCELUTIL", "Exception" + e.getMessage());
        }
        int[] size = new int[2];
        // Get the first sheet
        Sheet sheet = workbook.getSheet(0);
        int columns = sheet.getColumns();
        int rows = sheet.getRows();
        size[0] = columns;
        size[1] = rows;
        return size;
    }

    /**
     * 读取某一行
     *
     * @param row
     * @return
     */
    public ArrayList<String> ReadLineExcel(int row) {

        Workbook workbook = null;
        try {
            if (inputFile == null) {
                workbook = Workbook.getWorkbook(inputStream);
            } else {
                workbook = Workbook.getWorkbook(inputFile);
            }
        } catch (Exception e) {
            e.printStackTrace();

            Log.e("EXCELUTIL", "Exception" + e.getMessage());
        }
        ArrayList contents = new ArrayList<String[]>();
        // Get the first sheet
        Sheet sheet = workbook.getSheet(0);
        int columns = sheet.getColumns();

        // Loop over  column and lines
        for (int j = 0; j < columns; j++) {
            Cell cell = sheet.getCell(j, row);
            String string = cell.getContents();
            contents.add(string);
        }
        return contents;
    }

    /**
     * 读取某一列
     *
     * @param column
     * @return
     */
    public ArrayList<String> ReadColumnExcel(int column) {
        Workbook workbook = null;
        try {
            if (inputFile == null) {
                workbook = Workbook.getWorkbook(inputStream);
            } else {
                workbook = Workbook.getWorkbook(inputFile);
            }
        } catch (Exception e) {
            e.printStackTrace();
            Log.e("EXCELUTIL", "Exception" + e.getMessage());
        }
        ArrayList contents = new ArrayList<String[]>();
        // Get the first sheet
        Sheet sheet = workbook.getSheet(0);
        int rows = sheet.getRows();
        // Loop over  column and lines
        for (int i = 0; i < rows; i++) {
            Cell cell = sheet.getCell(column, i);
            String string = cell.getContents();
            contents.add(string);
        }

        return contents;
    }


    /**
     * 读取除首行外的内容行
     *
     * @return
     */
    public ArrayList ReadContentExcel() throws IOException, BiffException {
        int[] size = getExcelSize();
        return ReadExcelByColumn(0, size[0] - 1, 1, size[1] - 1);
    }

    public void WriteExcel(ArrayList<String[]> data) throws IOException, WriteException {
        WritableWorkbook workbook = Workbook.createWorkbook(outputFile, new WorkbookSettings());
        workbook.createSheet("New", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        createLabel(excelSheet);
        createStringContent(excelSheet, data);
        workbook.write();
        workbook.close();
    }

    private void createStringContent(WritableSheet sheet, ArrayList<String[]> data) throws WriteException,
            RowsExceededException {

        int columns = data.size();
        int rows;
        rows = data.get(0).length;
        // Loop over  column and lines
        for (int j = 0; j < columns; j++) {
            for (int i = 0; i < rows; i++) {
                addLabel(sheet, j, i, data.get(j)[i]);
            }
        }
    }

    private void createLabel(WritableSheet sheet)
            throws WriteException {
        // Lets create a times font
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
        // Define the cell format
        times = new WritableCellFormat(times10pt);

        // Lets automatically wrap the cells
        times.setWrap(true);

        // create create a bold font with unterlines
        WritableFont times10ptBoldUnderline = new WritableFont(WritableFont.TIMES, 10, WritableFont.BOLD, false,
                UnderlineStyle.SINGLE);
        timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
        // Lets automatically wrap the cells
        timesBoldUnderline.setWrap(true);

        CellView cv = new CellView();
        cv.setFormat(times);
        cv.setFormat(timesBoldUnderline);
        cv.setAutosize(true);

        // Write a few headers
        addCaption(sheet, 0, 0, "Header 1");
        addCaption(sheet, 1, 0, "This is another header");

    }


    private void addCaption(WritableSheet sheet, int column, int row, String s)
            throws RowsExceededException, WriteException {
        Label label;
        label = new Label(column, row, s, timesBoldUnderline);
        sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row,
                           Integer integer) throws WriteException, RowsExceededException {
        Number number;
        number = new Number(column, row, integer, times);
        sheet.addCell(number);
    }

    private void addLabel(WritableSheet sheet, int column, int row, String s)
            throws WriteException, RowsExceededException {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }
}
