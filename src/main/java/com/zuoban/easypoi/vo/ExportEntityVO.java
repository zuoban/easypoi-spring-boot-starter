package com.zuoban.easypoi.vo;

import java.util.List;

/**
 * 导出实体类型数据
 *
 * @author wangjinqiang
 * @date 2018-07-13
 */
public class ExportEntityVO extends BaseExportVO {
    /**
     * 实体类型
     */
    private Class dataClass;

    public Class getDataClass() {
        return dataClass;
    }

    public ExportEntityVO(List<?> data, Class dataClass, String fileName) {
        super(data, fileName);
        this.dataClass = dataClass;
    }
}
