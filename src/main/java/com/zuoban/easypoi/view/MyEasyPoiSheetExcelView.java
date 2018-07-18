package com.zuoban.easypoi.view;

import com.zuoban.easypoi.util.MyExcelExportUtil;
import com.zuoban.easypoi.vo.SheetExcelConstants;
import com.zuoban.easypoi.vo.SheetVO;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.servlet.view.AbstractView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

/**
 * @author wangjinqiang
 * @date 2018-07-05
 */
@Controller
public class MyEasyPoiSheetExcelView extends AbstractView {
    private static final String CONTENT_TYPE = "text/html;application/vnd.ms-excel";
    private static final String HSSF = ".xls";
    private static final String XSSF = ".xlsx";

    @Override
    public void setContentType(String contentType) {
        super.setContentType(CONTENT_TYPE);
    }

    @Override
    protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request, HttpServletResponse response) throws Exception {
        String codedFileName = "临时文件";

        SheetVO[] vos = (SheetVO[]) model.get(SheetExcelConstants.DATA);
        Workbook workbook = MyExcelExportUtil.exportExcel(vos);

        String fileName = vos[0].getExportVO().getFileName();

        if (StringUtils.isNotBlank(fileName)) {
            codedFileName = fileName;
        }
        if (workbook instanceof HSSFWorkbook) {
            codedFileName += HSSF;
        } else {
            codedFileName += XSSF;
        }
        if (isIE(request)) {
            codedFileName = java.net.URLEncoder.encode(codedFileName, "UTF8");
        } else {
            codedFileName = new String(codedFileName.getBytes("UTF-8"), "ISO-8859-1");
        }
        response.setHeader("content-disposition", "attachment;filename=" + codedFileName);
        ServletOutputStream out = response.getOutputStream();
        workbook.write(out);
        out.flush();
    }

    private static boolean isIE(HttpServletRequest request) {
        String userAgent = request.getHeader("USER-AGENT");

        if (userAgent == null) {
            return false;
        }

        return (userAgent.toLowerCase().indexOf("msie") > 0
                || userAgent.toLowerCase().indexOf("rv:11.0") > 0
                || userAgent.toLowerCase().indexOf("edge") > 0);

    }
}
