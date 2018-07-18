package com.zuoban.easypoi.util;

import com.zuoban.easypoi.vo.SheetExcelConstants;
import com.zuoban.easypoi.vo.SheetVO;
import org.springframework.util.Assert;
import org.springframework.web.servlet.ModelAndView;

/**
 * @author wangjinqiang
 * @date 2018-06-28
 */
public class MyExcelViewUtil {
    private MyExcelViewUtil() {
    }

    public static ModelAndView getSheetExcelModelAndView(SheetVO... vos) {
        Assert.isTrue(vos.length > 0, "vos can not be empty");
        ModelAndView modelAndView = new ModelAndView(SheetExcelConstants.VIEW_NAME);
        modelAndView.addObject(SheetExcelConstants.DATA, vos);
        return modelAndView;
    }
}
