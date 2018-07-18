package com.zuoban.easypoi.vo;

import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;

import java.util.List;

/**
 * 导出 map 类型数据封装
 *
 * @author wangjinqiang
 * @date 2018-06-28
 */
public class ExportMapVO extends BaseExportVO {
    /**
     * cell类型映射
     */
    private List<ExcelExportEntity> entityList;

    public List<ExcelExportEntity> getEntityList() {
        return entityList;
    }

    public ExportMapVO(List<ExcelExportEntity> entityList, List data, String fileName) {
        super(data, fileName);
        this.entityList = entityList;
    }
}
