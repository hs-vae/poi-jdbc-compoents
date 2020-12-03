package jdbc;

import Entity.ComponentsOne;
import Entity.Pathname;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.Scanner;

/**
 * 操作零件表在数据库中对应的数据类
 */

public class PartDB {


    public static void insertAllPart(ComponentsOne one) throws SQLException {
        String insertSql="insert into part(number,name,size)values(?,?,?)";
        DBUtil db=new DBUtil();
        db.getConnection();
        int a=0;
        for (int i = 1; i < one.number.length; i++) {
            Object[] param =new Object[]{one.getNumber()[i],one.getName()[i],one.getSize()[i]};
            a=db.executeUpdate(insertSql,param);
        }
        if(a!=0){
            System.out.println("零件表.xlsx信息导入数据库part表成功!");
        }
        db.closeAll();
        printPart();
    }

    public static void insertPart(ComponentsOne one, Pathname pn) throws IOException, SQLException {
        String insertSql = "insert into part(number,name,size)values(?,?,?)";
        DBUtil db = new DBUtil();
        db.getConnection();
        int a = 0;
        Scanner sc = new Scanner(System.in);
        System.out.println("请输入要插入的零件编号:");
        String number = sc.next();
        System.out.println("请输入要插入的零件名称:");
        String name = sc.next();
        System.out.println("请输入要插入的零件规格:");
        String size = sc.next();
        Object[] param = new Object[]{number, name, size};
        int b = db.executeUpdate(insertSql, param);
        if (b != 0) {
            System.out.println("插入成功数据到part表成功," + "已自动更新到" + pn.getPathnameOne() + "中!");
        }
        db.closeAll();
        printPart();

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("零件表");
        FileOutputStream fos = new FileOutputStream(new File(pn.getPathnameOne()));
        Workbook sheets = null;
        XSSFRow row = null;
        XSSFCell cell = null;
        //增加零件表每个零件信息数组的长度+1
        one.number = Arrays.copyOf(one.number, one.number.length + 1);
        one.name = Arrays.copyOf(one.name, one.name.length + 1);
        one.size = Arrays.copyOf(one.size, one.size.length + 1);
        for (int i = 0; i < one.number.length + 1; i++) { //4+1=5
            row = sheet.createRow(i);
            cell = row.createCell(i);
            if (i == 0) {
                //设置单元格范围地址,参数分别表示为(起始行号,终止行号,起始列号,终止列号)
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
                row.createCell(0).setCellValue("零件表信息");
                //设置水平居中
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cell.setCellStyle(style);
                sheet.addMergedRegion(region);
            } else {
                //在数组尾部添加新增的零件信息
                if (i == one.number.length) {
                    row.createCell(0).setCellValue(number);
                    row.createCell(1).setCellValue(name);
                    row.createCell(2).setCellValue(size);
                    one.number[i -1] = number;
                    one.name[i - 1] = name;
                    one.size[i - 1] = size;
                    one.setNumber(one.number);
                    one.setName(one.name);
                    one.setSize(one.size);
                } else {
                    row.createCell(0).setCellValue(one.getNumber()[i - 1]);
                    row.createCell(1).setCellValue(one.getName()[i - 1]);
                    row.createCell(2).setCellValue(one.getSize()[i - 1]);
                }
            }
        }
        workbook.write(fos);
        workbook.close();
        fos.close();
    }


    public static void selectPart(ComponentsOne one) throws SQLException {
        String selectSql="select number,name,size from part where number=?";
        DBUtil db=new DBUtil();
        db.getConnection();
        Scanner sc=new Scanner(System.in);
        System.out.println("输入你想查询的零件编号:");
        String number=sc.next();
        Object[] param=new Object[]{number};
        ResultSet rs=db.executeQuery(selectSql,param);
        while (rs.next()){
            System.out.println(rs.getString(1)+"       "+rs.getString(2)+"       "+rs.getString(3));
        }
        db.closeAll();
    }
    public static void updatePart(ComponentsOne one,Pathname pn) throws SQLException, IOException {
        String updateSql="update part set name=?,size=? where number=?";
        Scanner sc=new Scanner(System.in);
        System.out.println("输入你想修改的零件编号:");
        String number=sc.next();
        System.out.println("输入修改后的零件名称:");
        String name=sc.next();
        System.out.println("输入修改后的零件规格:");
        String size=sc.next();
        //将更新的零件信息也更新到零件想信息实体类
        for (int i = 0; i < one.number.length; i++) {
            if(number.equals(one.getNumber()[i])){
                one.number[i]=number;
                one.name[i]=name;
                one.size[i]=size;
                one.setNumber(one.number);
                one.setName(one.name);
                one.setSize(one.size);
            }
        }
        DBUtil db=new DBUtil();
        db.getConnection();
        Object[] param=new Object[]{name,size,number};
        int b=db.executeUpdate(updateSql,param);
        if (b!=0){
            System.out.println("从part表修改该条零件信息成功!"+"已自动更新到"+pn.getPathnameOne()+"中!");
        }
        db.closeAll();
        printPart();


        FileInputStream fis=new FileInputStream(pn.getPathnameOne());
        XSSFWorkbook workbook = new XSSFWorkbook();
        //输入编号打印零件名称和零件规格
        Workbook sheets1=new XSSFWorkbook(fis);
        Sheet sheet1 = sheets1.getSheetAt(0);
        XSSFRow row = null;
        XSSFCell cell = null;
        FileOutputStream fos=null;
        XSSFSheet sheet = workbook.createSheet("零件表");
        fos = new FileOutputStream(pn.getPathnameOne());
        //设置标题
        row = sheet.createRow(0);
        cell = row.createCell(0);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
        row.createCell(0).setCellValue("零件表信息");
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellStyle(style);
        sheet.addMergedRegion(region);
        for (int i = 0; i < one.getNumber().length; i++) {
            row = sheet.createRow(i);
            cell = row.createCell(i);
            row.createCell(0).setCellValue(one.getNumber()[i]);
            row.createCell(1).setCellValue(one.getName()[i]);
            row.createCell(2).setCellValue(one.getSize()[i]);
        }
        workbook.write(fos);
        workbook.close();
        fos.close();
        //打印输出修改后的零件表信息
        fis.close();
        sheets1.close();

    }
    public static void deletePrat(ComponentsOne one,Pathname pn) throws SQLException, IOException {
        String deleteSql="delete from part where number=?";;
        Scanner sc=new Scanner(System.in);
        System.out.println("输入删除的零件对应的零件编号:");
        String number = sc.next();
        DBUtil db=new DBUtil();
        db.getConnection();
        Object[] param=new Object[]{number};
        int c=db.executeUpdate(deleteSql,param);
        if (c!=0){
            System.out.println("从part表删除该条零件信息成功!"+"已自动更新到"+pn.getPathnameOne()+"中!");
        }
        db.closeAll();
        printPart();

        int b=0;
        for (int i = 0; i < one.number.length; i++) {
            if(number.equals(one.getNumber()[i])){
                b=i;
            }
        }
        for (int i = b; i < one.number.length-1; i++) { //1 2
            one.number[i]=one.getNumber()[i+1];
            one.name[i]=one.getName()[i+1];
            one.size[i]=one.getSize()[i+1];
            one.setNumber(one.number);
            one.setName(one.name);
            one.setSize(one.size);
        }
        //将零件信息的数组长度缩小一个长度
        one.number=Arrays.copyOf(one.number,one.number.length-1);
        one.name=Arrays.copyOf(one.name,one.name.length-1);
        one.size=Arrays.copyOf(one.size,one.size.length-1);

        //创建一个删除指定零件编号的零件表
        FileOutputStream fos=new FileOutputStream(new File(pn.getPathnameOne()));
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("零件表");
        XSSFRow rowDelete=null;
        XSSFCell cellDelete=null;
        for (int i = 0; i < one.number.length+1; i++) { // 0 1 2 3
            rowDelete=sheet.createRow(i);
            cellDelete=rowDelete.createCell(i);
            if(i==0){
                rowDelete = sheet.createRow(0);
                cellDelete = rowDelete.createCell(0);
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 2);
                rowDelete.createCell(0).setCellValue("零件表信息");
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cellDelete.setCellStyle(style);
                sheet.addMergedRegion(region);
            }else{
                rowDelete.createCell(0).setCellValue(one.getNumber()[i-1]);
                rowDelete.createCell(1).setCellValue(one.getName()[i-1]);
                rowDelete.createCell(2).setCellValue(one.getSize()[i-1]);
            }
        }
        workbook.write(fos);
        workbook.close();
        fos.close();
    }
    public static void printPart() throws SQLException {
        DBUtil db=new DBUtil();
        db.getConnection();
        String sql="select * from part";
        ResultSet rs=db.executeQuery(sql,null);
        System.out.println("零件编号  零件名称  零件规格");
        while (rs.next()){
            System.out.println(rs.getString(1)+"       "+rs.getString(2)+"       "+rs.getString(3));
        }
        db.closeAll();
    }
}
