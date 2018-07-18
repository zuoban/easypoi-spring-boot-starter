package com.zuoban.easypoi.vo;

import java.util.List;

/**
 * easypoi 实体类型数据封装
 *
 * @author wangjinqiang
 * @date 2018-06-28
 */
public abstract class BaseExportVO {
    /**
     * 数据
     */
    private List<?> data;
    /**
     * 导出文件名
     */
    private String fileName;

    public List<?> getData() {
        return data;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public BaseExportVO(List<?> data, String fileName) {
        this.data = data;
        this.fileName = fileName;
    }
}
