package com.zuoban.easypoi.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;

/**
 * @author wangjinqiang
 * @date 2018-07-18
 */
public class Product0 {
    @Excel(name = "id")
    private Integer id;
    @Excel(name = "name")
    private String name;

    public Product0(Integer id, String name) {
        this.id = id;
        this.name = name;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
