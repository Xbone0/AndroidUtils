package org.hahaqu.app.lockbankaccount.operator;

import android.content.Context;
import android.content.res.AssetManager;

import org.hahaqu.app.lockbankaccount.util.ExcelUtil;

import java.io.IOException;
import java.util.ArrayList;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;

/**
 * Created by crossbones on 16/6/6.
 */
public class ExcelOperator {
    private Context context;
    private ExcelUtil excelUtil = null;
    private String assets;

    public ExcelOperator(Context context, String assets) {
        this.context = context;
        this.assets = assets;
        this.excelUtil = ExcelUtil.getInstance();
    }

    private void setInputAssets() throws IOException {
        AssetManager manager = context.getAssets();
        excelUtil.inputStream = manager.open(assets);
    }

    public ArrayList<String[]> ReadExcelByColumn() throws IOException {
        setInputAssets();
        return excelUtil.ReadExcelByColumn();
    }

    public ArrayList<String[]> ReadExcelByColumn(int column_start, int column_end, int row_start, int row_end) throws IOException {
        setInputAssets();
        return excelUtil.ReadExcelByColumn(column_start, column_end, row_start, row_end);
    }

    public ArrayList<String[]> ReadExcelByRow(int column_start, int column_end, int row_start, int row_end) throws IOException {
        setInputAssets();
        return excelUtil.ReadExcelByRow(column_start, column_end, row_start, row_end);
    }

    /**
     * 返回行列数
     *
     * @return
     */
    public int[] getExcelSize() throws IOException {
        setInputAssets();
        return excelUtil.getExcelSize();
    }

    /**
     * 读取某一行
     *
     * @param row
     * @return
     */
    public ArrayList<String> ReadLineExcel(int row) throws IOException {
        setInputAssets();
        return excelUtil.ReadLineExcel(row);
    }

    /**
     * 读取某一列
     *
     * @param column
     * @return
     */
    public ArrayList<String> ReadColumnExcel(int column) throws IOException {
        setInputAssets();
        return excelUtil.ReadColumnExcel(column);
    }


    /**
     * 按列读取除首行外的内容行
     *
     * @return
     */
    public ArrayList ReadContentByColumn() throws IOException, BiffException {
        int[] size = getExcelSize();
        return ReadExcelByColumn(0, size[0] - 1, 1, size[1] - 1);
    }

    /**
     * 按行读取除首行外的内容行
     *
     * @return
     */
    public ArrayList ReadContentByRow() throws IOException, BiffException {
        int[] size = getExcelSize();
        return ReadExcelByRow(0, size[0] - 1, 1, size[1] - 1);
    }

    public void WriteExcel(ArrayList<String[]> data) throws IOException, WriteException {
        excelUtil.WriteExcel(data);
    }

}
