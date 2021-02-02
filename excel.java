package excelDivide;

/**
 * @PackageName: excelDivide
 * @ClassName: Excel
 * @Description: //TODO
 * @Author: LEIBF
 * @Date: 2020/11/20 20:45
 */

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Excel {

    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub
        List<List<Object>> list = new ArrayList<List<Object>>();
        //原文件
        String filePath = "C:\\Users\\Zhangh\\Desktop\\1111.xlsx";
        //读取文件的sheet，改第一个数字
        Sheet sheet = new Sheet(4, 1);
        FileInputStream in = new FileInputStream(new File(filePath));
        List<Object> read = EasyExcelFactory.read(in, sheet);
        for (int i = 0; i < read.size(); i++) {

            String readStr = read.get(i).toString();
            String[] split = readStr.split(",");
            //项目
            String project = split[0].substring(1);
            //楼栋
            String building = split[1];
            //单元
            String cell = split[2];
            //房号
            String room = split[3];
            //名字
            String name = split[4];
            //电话
            String phone = split[5];
            //身份证
            String id = split[6];
            //住址
            String add = split[7].substring(0, split[7].length() - 1);
            if (name.contains(";")) {
                String[] names = name.split(";");
                //名字有两个
                String name1 = names[0].trim();
                String name2 = names[1].trim();
                String[] phones = phone.split(";");
                //电话两个
                String phone1 = phones[0].trim();
                String phone2 = phones[1].trim();
                String[] ids = id.split(";");
                //身份证两个
                String id1 = ids[0].trim();
                String id2 = ids[1].trim();
                String[] adds = add.split(";");
                //地址两个
                String add1 = adds[0].trim();
                String add2 = adds[1].trim();
                simpleWrite(list, project, building, cell, room, name1, phone1, id1, add2);
                simpleWrite(list, project, building, cell, room, name2, phone2, id2, add2);
            } else {
                simpleWrite(list, project, building, cell, room, name.trim(), phone.trim(), id.trim(), add.trim());
            }
        }
    }

    public static void simpleWrite(List<List<Object>> list, String project, String building, String cell, String room, String name, String phone, String id, String add) {
//		System.out.println(name);
        // 文件输出位置
        String outPath = "C:\\Users\\Zhangh\\Desktop\\test2.xlsx";

        try {
            // 所有行的集合
//			for (int i = 1; i <= 150; i++) {
            // 第 n 行的数据
            List<Object> row = new ArrayList<Object>();
            row.add(project);
            row.add(building);
            row.add(cell);
            row.add(room);
            row.add(name);
            row.add(phone);
            row.add(id);
            row.add(add);
            list.add(row);
//			}

            ExcelWriter excelWriter = EasyExcelFactory.getWriter(new FileOutputStream(outPath));
            // 表单
            Sheet sheet = new Sheet(1, 0);
            sheet.setSheetName("Sheet1");
            // 创建一个表格
            Table table = new Table(1);

            excelWriter.write1(list, sheet, table);
            // 记得 释放资源
            excelWriter.finish();
            System.out.println("ok");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

    }
}
