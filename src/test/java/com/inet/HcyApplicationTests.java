package com.inet;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import com.inet.codebase.entity.User;
import com.inet.codebase.service.UserService;
import org.apache.coyote.http11.InputFilter;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import javax.annotation.Resource;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

@SpringBootTest
class HcyApplicationTests {

    @Resource
    private UserService userService;

    /**
     * 查询全部
     * @author HCY
     * @since 2020-10-14
     */
    @Test
    void contextLoads() {
        List<User> list = userService.list();

        list.forEach(System.out::println);

    }

    /**
     * 导入Excel
     * @author HCY
     * @since 2020-10-14
     * @throws Exception
     */
    @Test
    void contextLoads1() throws Exception {
        File file = new File("C:\\Users\\Administrator.DESKTOP-TSJVEJ5\\Desktop\\test\\test.xlsx");

        ImportParams params = new ImportParams();
        //设置标题
        params.setTitleRows(1);
        //设置说明
        params.setHeadRows(1);
        //导入获取集合
        List<User> users = ExcelImportUtil.importExcel(file, User.class, params);
        //遍历集合
        users.forEach(System.out :: println);
        //进行批量添加
        boolean batch = userService.saveBatch(users);
        //输出结果
        System.out.println(batch);
    }

    /**
     * 导出Excel
     * @author HCY
     * @since 2020-10-14
     */
    @Test
    void contextLoads2() throws Exception {
        //查询全部
        List<User> list = userService.list();
        //设置Excel的描述文件
        ExportParams exportParams = new ExportParams("用户列表的所有数据", "用户信息" , ExcelType.XSSF);
        //进行导出的基本操作
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, User.class, list);
        //输入输出流地址
        FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\Administrator.DESKTOP-TSJVEJ5\\Desktop\\test\\users.xlsx");
        //进行输出流
        workbook.write(fileOutputStream);
        //关流
        fileOutputStream.close();
        workbook.close();
    }

    @Test
    void contextLoads3() {
    }


}
