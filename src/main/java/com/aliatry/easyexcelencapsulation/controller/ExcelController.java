package com.aliatry.easyexcelencapsulation.controller;

import com.aliatry.easyexcelencapsulation.domain.ImportInfo;
import com.aliatry.easyexcelencapsulation.tdo.ExportInfo;
import com.aliatry.easyexcelencapsulation.util.ExcelUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel控制层
 *
 * @author Simon
 */
@RestController
public class ExcelController {

    /**
     * 读取 Excel（指定某个 sheet）
     *
     * @param excel       文件
     * @param sheetNo     sheet 的序号 从1开始
     * @param headLineNum 表头行数，默认为1
     */
    @RequestMapping(value = "readExcel", method = RequestMethod.POST)
    public Object readExcel(@RequestParam(required = true) MultipartFile excel, @RequestParam(defaultValue = "1") int sheetNo,
                            @RequestParam(defaultValue = "1") int headLineNum) {
        return ExcelUtil.readExcel(excel, new ImportInfo(), sheetNo, headLineNum);
    }

    /**
     * 导出 Excel（一个 sheet）
     */
    @RequestMapping(value = "writeExcel", method = RequestMethod.GET)
    public void writeExcel(HttpServletResponse response) {
        List<ExportInfo> list = getList();
        String fileName = "ExcelFile";
        String sheetName = "one sheet";

        ExcelUtil.writeExcel(response, list, fileName, sheetName, new ExportInfo());
    }

    /**
     * 读取 Excel（允许多个 sheet）
     *
     * @param excel 文件
     */
    @RequestMapping(value = "readExcelWithSheets", method = RequestMethod.POST)
    public Object readExcelWithSheets(@RequestParam(required = true) MultipartFile excel) {
        return ExcelUtil.readExcel(excel, new ImportInfo());
    }

    /**
     * 导出 Excel（多个 sheet）
     */
    @RequestMapping(value = "writeExcelWithSheets", method = RequestMethod.GET)
    public void writeExcelWithSheets(HttpServletResponse response) {
        List<ExportInfo> list = getList();
        String fileName = "ExcelFile";
        String sheetName1 = "one sheet";
        String sheetName2 = "two sheet";
        String sheetName3 = "three sheet";

        ExcelUtil.writeExcelWithSheets(response, list, fileName, sheetName1, new ExportInfo())
                .write(list, sheetName2, new ExportInfo())
                .write(list, sheetName3, new ExportInfo())
                .finish();
    }

    private List<ExportInfo> getList() {
        List<ExportInfo> list = new ArrayList<>();
        ExportInfo model1 = new ExportInfo();
        model1.setName("admin");
        model1.setAge("20");
        model1.setEmail("admin_simon@icloud.com");
        list.add(model1);
        ExportInfo model2 = new ExportInfo();
        model2.setName("simon");
        model2.setAge("21");
        model2.setEmail("admin_simon@icloud.com");
        list.add(model2);
        return list;
    }
}