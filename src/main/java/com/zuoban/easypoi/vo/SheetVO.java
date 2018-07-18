package com.zuoban.easypoi.vo;

import cn.afterturn.easypoi.excel.entity.ExportParams;

/**
 * sheet 数据封装
 *
 * @author wangjinqiang
 * @date 2018-07-04
 */
public class SheetVO {
    /**
     * 导入参数
     */
    private ExportParams exportParams;
    /**
     * 导出数据
     */
    private BaseExportVO exportVO;

    public SheetVO(ExportParams exportParams, BaseExportVO exportVO) {
        this.exportParams = exportParams;
        this.exportVO = exportVO;
    }

    public ExportParams getExportParams() {
        return exportParams;
    }

    public void setExportParams(ExportParams exportParams) {
        this.exportParams = exportParams;
    }

    public BaseExportVO getExportVO() {
        return exportVO;
    }

    public void setExportVO(BaseExportVO exportVO) {
        this.exportVO = exportVO;
    }

    public static SheetVO of(ExportParams params, BaseExportVO vo) {
        return new SheetVO(params, vo);
    }
}
