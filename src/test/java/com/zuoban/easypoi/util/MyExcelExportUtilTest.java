package com.zuoban.easypoi.util;

import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.zuoban.easypoi.entity.Product0;
import com.zuoban.easypoi.vo.BaseExportVO;
import com.zuoban.easypoi.vo.ExportEntityVO;
import com.zuoban.easypoi.vo.ExportMapVO;
import com.zuoban.easypoi.vo.SheetVO;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * @author wangjinqiang
 * @date 2018-07-18
 */
public class MyExcelExportUtilTest {
    private List<Product0> productList1;
    private List<Product0> productList2;

    private List<Map> mapList1;
    private List<Map> mapList2;
    private List<ExcelExportEntity> entityList1;
    private List<ExcelExportEntity> entityList2;

    @Before
    public void init() {
        productList1 = new ArrayList<>();
        productList2 = Lists.newArrayList();
        mapList1 = Lists.newArrayList();
        mapList2 = Lists.newArrayList();
        entityList1 = Lists.newArrayList();
        entityList2 = Lists.newArrayList();

        for (int i = 0; i < 100; i++) {
            Product0 temp = new Product0(i, "test1" + i);
            productList1.add(temp);
            productList2.add(temp);
            Map<String, Object> map1 = Maps.newHashMap();
            map1.put("map1-col1", i);
            map1.put("map1-col2", RandomStringUtils.randomAlphabetic(4));
            map1.put("map1-col3", RandomStringUtils.randomAlphabetic(4));
            mapList1.add(map1);


            Map<String, Object> map2 = Maps.newHashMap();
            map2.put("map2-col1", i);
            map2.put("map2-col2", RandomStringUtils.randomAlphabetic(4));
            map2.put("map2-col3", RandomStringUtils.randomAlphabetic(4));
            map2.put("map2-col4", RandomStringUtils.randomAlphabetic(4));
            mapList2.add(map2);
        }

        entityList1.add(new ExcelExportEntity("第一列", "map1-col1"));
        entityList1.add(new ExcelExportEntity("第二列", "map1-col2"));
        entityList1.add(new ExcelExportEntity("第三列", "map1-col3"));

        entityList2.add(new ExcelExportEntity("第一列", "map2-col1"));
        entityList2.add(new ExcelExportEntity("第二列", "map2-col2"));
        entityList2.add(new ExcelExportEntity("第三列", "map2-col3"));
        entityList2.add(new ExcelExportEntity("第四列", "map2-col4"));
    }


    /**
     * 导出实体数据的单sheet文件
     */
    @Test
    public void exportSingleSheetExcelWithEntity() {
        // 组装 sheet 数据
        ExportParams exportParams = new ExportParams("sheet1 title", "sheet1");
        ExportEntityVO exportVO = new ExportEntityVO(productList1, Product0.class, "exportSingleSheetExcelWithEntity");
        SheetVO sheetVO = new SheetVO(exportParams, exportVO);
        Workbook workbook = MyExcelExportUtil.exportExcel(sheetVO);
        writeToFile(workbook, sheetVO);
    }

    @Test
    public void exportSingleSheetExcelWithMap() {
        // sheet1 data
        ExportParams exportParams1 = new ExportParams("map sheet1 title", "sheet1");
        ExportMapVO vo1 = new ExportMapVO(entityList1, mapList1, "exportSingleSheetExcelWithMap");
        SheetVO sheetVO = new SheetVO(exportParams1, vo1);
        Workbook workbook = MyExcelExportUtil.exportExcel(sheetVO);
        writeToFile(workbook, sheetVO);
    }

    private void writeToFile(Workbook workbook, SheetVO firstSheet) {
        BaseExportVO exportVO = firstSheet.getExportVO();
        ExportParams exportParams = firstSheet.getExportParams();

        String suffix = exportParams.getType().equals(ExcelType.HSSF) ? ".xls" : ".xlsx";
        String fullName = "export/" + exportVO.getFileName() + suffix;

        // 输出
        try {
            workbook.write(new FileOutputStream(fullName));
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 实体数据生成多 sheet 测试
     */
    @Test
    public void exportExcel() {
        // 组装 sheet 数据
        ExportParams exportParams = new ExportParams("sheet1 title", "sheet1");
        ExportEntityVO exportVO = new ExportEntityVO(productList1, Product0.class, "many-sheet");
        SheetVO sheetVO1 = new SheetVO(exportParams, exportVO);

        // 组装 sheet 数据
        ExportParams exportParams2 = new ExportParams("sheet2 title", "sheet2");
        ExportEntityVO exportVO2 = new ExportEntityVO(productList2, Product0.class, "many-sheet");
        SheetVO sheetVO2 = new SheetVO(exportParams2, exportVO2);


        // 调用工具类生成 workbook
        Workbook workbook = MyExcelExportUtil.exportExcel(sheetVO1, sheetVO2);
        writeToFile(workbook, sheetVO1);
    }

    /**
     * map 数据多 sheet 导出测试
     */
    @Test
    public void exportExcelWithMap() {
        // sheet1 data
        ExportParams exportParams1 = new ExportParams("map sheet1 title", "sheet1");
        ExportMapVO vo1 = new ExportMapVO(entityList1, mapList1, "exportExcelWithMap");

        // sheet2 data
        ExportParams exportParams2 = new ExportParams("map sheet2 title", "sheet2");
        ExportMapVO vo2 = new ExportMapVO(entityList2, mapList2, null);

        SheetVO sheetVO1 = new SheetVO(exportParams1, vo1);
        SheetVO sheetVO2 = new SheetVO(exportParams2, vo2);

        Workbook workbook = MyExcelExportUtil.exportExcel(sheetVO1, sheetVO2);

        writeToFile(workbook, sheetVO1);
    }

    /**
     * 实体和 map 混合 生成多 sheet 测试
     * <p>
     * throws IOException
     */
    @Test
    public void exportExcelMix() {

        // map data
        ExportParams exportParams1 = new ExportParams("map sheet1 title", "sheet1");
        ExportMapVO vo1 = new ExportMapVO(entityList1, mapList1, "exportExcelMix");
        SheetVO sheetVO1 = new SheetVO(exportParams1, vo1);

        // entity data
        ExportParams exportParams2 = new ExportParams("sheet2 title", "sheet2");
        ExportEntityVO vo2 = new ExportEntityVO(productList2, Product0.class, null);
        SheetVO sheetVO2 = new SheetVO(exportParams2, vo2);

        Workbook workbook = MyExcelExportUtil.exportExcel(sheetVO1, sheetVO2);
        writeToFile(workbook, sheetVO1);
    }


    @Test
    public void emptyMapData() {
        ExportParams exportParams = new ExportParams("map sheet1 title", "sheet1");
        ExportMapVO vo = new ExportMapVO(entityList1, Collections.emptyList(), "emptyMapData");
        vo.setFileName("testEmptyData");
        SheetVO sheetVO = new SheetVO(exportParams, vo);

        Workbook workbook = MyExcelExportUtil.exportExcel(sheetVO);
        writeToFile(workbook, sheetVO);
    }

    @Test
    public void emptyEntityData() {
        // entity data
        ExportParams exportParams2 = new ExportParams("sheet2 title", "sheet2");
        ArrayList<Product0> product0s = Lists.newArrayList();

        ExportEntityVO vo2 = new ExportEntityVO(product0s, Product0.class, "emptyEntityData");
        SheetVO sheetVO = new SheetVO(exportParams2, vo2);

        Workbook workbook = MyExcelExportUtil.exportExcel(sheetVO);
        writeToFile(workbook, sheetVO);
    }
}