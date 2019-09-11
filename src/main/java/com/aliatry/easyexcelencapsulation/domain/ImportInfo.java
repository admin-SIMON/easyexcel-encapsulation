package com.aliatry.easyexcelencapsulation.domain;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * 导入 Excel 时使用的映射实体类,Excel 模型
 * @author Simon
 */
public class ImportInfo extends BaseRowModel {
    @ExcelProperty(index = 0)
    private String name;

    @ExcelProperty(index = 1)
    private String age;

    @ExcelProperty(index = 2)
    private String email;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

}