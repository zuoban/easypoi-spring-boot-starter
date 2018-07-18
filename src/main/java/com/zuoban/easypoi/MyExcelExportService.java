package com.zuoban.easypoi;

import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.entity.vo.BaseEntityTypeConstants;
import cn.afterturn.easypoi.excel.entity.vo.PoiBaseConstants;
import cn.afterturn.easypoi.excel.export.ExcelExportService;
import cn.afterturn.easypoi.excel.export.base.ExportCommonService;
import cn.afterturn.easypoi.exception.excel.ExcelExportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelExportEnum;
import com.zuoban.easypoi.vo.BaseExportVO;
import com.zuoban.easypoi.vo.ExportEntityVO;
import com.zuoban.easypoi.vo.ExportMapVO;
import com.zuoban.easypoi.vo.SheetVO;
import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collection;
import java.util.List;

/**
 * @author wangjinqiang
 * @date 2018-07-02
 */
public class MyExcelExportService extends ExcelExportService {

    private int currentIndex = 0;

    @Override
    public int createCells(Drawing patriarch, int index, Object t,
                           List<ExcelExportEntity> excelParams, Sheet sheet, Workbook workbook,
                           short rowHeight) {
        try {
            ExcelExportEntity entity;
            Row row = sheet.createRow(index);
            if (rowHeight != -1) {
                row.setHeight(rowHeight);
            }
            int maxHeight = 1, cellNum = 0;
            int indexKey = createIndexCell(row, index, excelParams.get(0));
            cellNum += indexKey;
            for (int k = indexKey, paramSize = excelParams.size(); k < paramSize; k++) {
                entity = excelParams.get(k);
                if (entity.getList() != null) {
                    Collection<?> list = getListCellValue(entity, t);
                    int listC = 0;
                    if (list != null && list.size() > 0) {
                        for (Object obj : list) {
                            createListCells(patriarch, index + listC, cellNum, obj, entity.getList(),
                                    sheet, workbook, rowHeight);
                            listC++;
                        }
                    }
                    cellNum += entity.getList().size();
                    if (list != null && list.size() > maxHeight) {
                        maxHeight = list.size();
                    }
                } else {
                    Object value = getCellValue(entity, t);


                    if (entity.getType() == BaseEntityTypeConstants.STRING_TYPE) {
                        createStringCell(row, cellNum++, value == null ? "" : value.toString(),
                                getStyles(row.getRowNum(), entity, t, value),
                                entity);
                        if (entity.isHyperlink()) {
                            row.getCell(cellNum - 1)
                                    .setHyperlink(dataHandler.getHyperlink(
                                            row.getSheet().getWorkbook().getCreationHelper(), t,
                                            entity.getName(), value));
                        }
                    } else if (entity.getType() == BaseEntityTypeConstants.DOUBLE_TYPE) {
                        createDoubleCell(row, cellNum++, value == null ? "" : value.toString(),
                                index % 2 == 0 ? getStyles(false, entity) : getStyles(true, entity),
                                entity);
                        if (entity.isHyperlink()) {
                            row.getCell(cellNum - 1).setHyperlink(dataHandler.getHyperlink(
                                    row.getSheet().getWorkbook().getCreationHelper(), t,
                                    entity.getName(), value));
                        }
                    } else {
                        createImageCell(patriarch, entity, row, cellNum++,
                                value == null ? "" : value.toString(), t);
                    }
                }
            }
            // 合并需要合并的单元格
            cellNum = 0;
            for (int k = indexKey, paramSize = excelParams.size(); k < paramSize; k++) {
                entity = excelParams.get(k);
                if (entity.getList() != null) {
                    cellNum += entity.getList().size();
                } else if (entity.isNeedMerge() && maxHeight > 1) {
                    for (int i = index + 1; i < index + maxHeight; i++) {
                        sheet.getRow(i).createCell(cellNum);
                        sheet.getRow(i).getCell(cellNum).setCellStyle(getStyles(false, entity));
                    }
                    sheet.addMergedRegion(
                            new CellRangeAddress(index, index + maxHeight - 1, cellNum, cellNum));
                    cellNum++;
                }
            }
            return maxHeight;
        } catch (Exception e) {
            ExportCommonService.LOGGER.error("excel cell export error ,data is :{}", ReflectionToStringBuilder.toString(t));
            throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
        }

    }

    /**
     * 创建多 Sheet excel
     *
     * @param workbook
     * @param voList
     */
    public void createSheets(Workbook workbook, List<SheetVO> voList) {
        for (SheetVO sheetVO : voList) {
            createSheet(workbook, sheetVO);
        }
    }

    private int createIndexCell(Row row, int index, ExcelExportEntity excelExportEntity) {
        if (excelExportEntity.getName() != null && "序号".equals(excelExportEntity.getName()) && excelExportEntity.getFormat() != null
                && excelExportEntity.getFormat().equals(PoiBaseConstants.IS_ADD_INDEX)) {
            createStringCell(row, 0, currentIndex + "",
                    index % 2 == 0 ? getStyles(false, null) : getStyles(true, null), null);

            currentIndex = currentIndex + 1;
            return 1;
        }
        return 0;
    }

    private void createSheet(Workbook workbook, SheetVO vo) {
        BaseExportVO exportVO = vo.getExportVO();
        if (exportVO instanceof ExportMapVO) {
            // map
            createSheetForMap(workbook, vo.getExportParams(), ((ExportMapVO) exportVO).getEntityList(), exportVO.getData());
        } else if (exportVO instanceof ExportEntityVO) {
            // entity
            createSheet(workbook, vo.getExportParams(), ((ExportEntityVO) exportVO).getDataClass(), exportVO.getData());
        }
    }

    private CellStyle getStyles(int dataRow, ExcelExportEntity entity, Object obj, Object data) {
        return excelExportStyler.getStyles(null, dataRow, entity, obj, data);
    }
}

