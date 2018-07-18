package com.zuoban.easypoi.util;

import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import com.zuoban.easypoi.MyExcelExportService;
import com.zuoban.easypoi.vo.SheetVO;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Arrays;
import java.util.List;

/**
 * @author wangjinqiang
 * @date 2018-07-02
 */
public class MyExcelExportUtil {
    private MyExcelExportUtil() {
    }

    public static Workbook getWorkbook(ExcelType type, int size) {
        if (ExcelType.HSSF.equals(type)) {
            return new HSSFWorkbook();
        } else if (size < 100000) {
            return new XSSFWorkbook();
        } else {
            return new SXSSFWorkbook();
        }
    }

    /**
     * 导出
     *
     * @param vos
     * @return
     */
    public static Workbook exportExcel(SheetVO... vos) {

        // 创建 Workbook
        SheetVO first = vos[0];
        ExportParams exportParams = first.getExportParams();
        // 若 exportParams 为空， 则创建一个
        if (exportParams == null) {
            exportParams = new ExportParams();
        }


        long totalSize = Arrays.stream(vos).mapToLong(it -> {
            List data = it.getExportVO().getData();
            return data == null ? 0 : data.size();
        }).sum();
        Workbook workbook = getWorkbook(exportParams.getType(), Math.toIntExact(totalSize));

        // 填充 sheet 数据
        new MyExcelExportService().createSheets(workbook, Arrays.asList(vos));
        return workbook;
    }

}
