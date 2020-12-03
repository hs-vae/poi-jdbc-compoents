package jdbc;

import Entity.ComponentsTwo;
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
 * 操作零件进库表在数据库中对应的数据类
 */
public class PartEntryDB {


    /**
     *  将零件进库表的零件信息导入数据库compoents的partentry表中
     */
    public static void insertAllPartEntry(ComponentsTwo two) throws SQLException {
        String insertSql="insert into partentry(number,amount)values(?,?)";
        DBUtil db=new DBUtil();
        db.getConnection();
        int a=0;
        for (int i = 1; i < two.number.length; i++) {
            Object[] param =new Object[]{two.getNumber()[i],two.getAmount()[i]};
            a=db.executeUpdate(insertSql,param);
        }
        if(a!=0){
            System.out.println("零件进库表.xlsx信息导入数据库partentry表成功!");
        }
        db.closeAll();
        printPartEntry();
    }


    public static void insertPartEntry(ComponentsTwo two, Pathname pn) throws SQLException, IOException {
        String insertSql = "insert into partentry(number,amount)values(?,?)";
        DBUtil db = new DBUtil();
        db.getConnection();
        Scanner sc = new Scanner(System.in);
        System.out.println("请输入要插入的零件编号:");
        String number = sc.next();
        System.out.println("请输入要插入的零件名称:");
        int amount = sc.nextInt();

        Object[] param = new Object[]{number,amount};
        int b = db.executeUpdate(insertSql, param);
        if (b != 0) {
            System.out.println("插入成功数据到partentry表成功," + "已自动更新到" + pn.getPathnameTwo() + "中!");
        }
        db.closeAll();
        printPartEntry();

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("零件进库表");
        FileOutputStream fos=new FileOutputStream(new File(pn.getPathnameTwo()));
        Workbook sheets=null;
        XSSFRow row=null;
        XSSFCell cell=null;
        //增加零件表每个零件信息数组的长度+1
        two.number = Arrays.copyOf(two.number,two.number.length+1);
        two.amount = Arrays.copyOf(two.amount,two.amount.length+1);

        for (int i = 0; i < two.number.length+1; i++) { //4+1=5
            row = sheet.createRow(i);
            cell = row.createCell(i);
            if (i == 0) {
                //设置单元格范围地址,参数分别表示为(起始行号,终止行号,起始列号,终止列号)
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 1);
                row.createCell(0).setCellValue("零件进库表信息");
                //设置水平居中
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cell.setCellStyle(style);
                sheet.addMergedRegion(region);
            }  else {
                //在数组尾部添加新增的零件信息
                if(i==two.number.length){
                    row.createCell(0).setCellValue(number);
                    row.createCell(1).setCellValue(amount);
                    two.number[i-1]=number;
                    two.amount[i-1]=amount;
                    two.setNumber(two.number);
                    two.setAmount(two.amount);
                }else if(i==1){
                    row.createCell(0).setCellValue(two.getNumber()[i-1]);
                    row.createCell(1).setCellValue("零件数量");
                }else{
                    row.createCell(0).setCellValue(two.getNumber()[i-1]);
                    row.createCell(1).setCellValue(two.getAmount()[i-1]);
                }
            }
        }
        workbook.write(fos);
        workbook.close();
        fos.close();

    }
    public static void selectPartEntry(ComponentsTwo two) throws SQLException {
        String selectSql="select number,amount from partentry where number=?";
        DBUtil db=new DBUtil();
        db.getConnection();
        System.out.println("输入想要查询的零件编号:");
        Scanner sc=new Scanner(System.in);
        String number =sc.next();
        Object[] param=new Object[]{number};
        ResultSet rs=db.executeQuery(selectSql,param);
        System.out.println("零件编号"+"   "+"零件数量");
        while (rs.next()){
            System.out.println(rs.getString(1)+"       "+rs.getInt(2));
        }
        db.closeAll();
    }

    public static void updatePartEntry(ComponentsTwo two,Pathname pn) throws SQLException, IOException {
        String updateSql="update partentry set amount=? where number=?";
        Scanner sc=new Scanner(System.in);
        System.out.println("输入你想修改的零件编号:");
        String number=sc.next();
        System.out.println("输入修改后的零件数量:");
        int amount=sc.nextInt();
        //将更新的零件信息也更新到零件想信息实体类
        for (int i = 0; i < two.number.length; i++) {
            if(number.equals(two.getNumber()[i])){
                two.number[i]=number;
                two.amount[i]=amount;
                two.setNumber(two.number);
                two.setAmount(two.amount);
            }
        }
        DBUtil db=new DBUtil();
        db.getConnection();
        Object[] param=new Object[]{amount,number};
        int c=db.executeUpdate(updateSql,param);
        if (c!=0){
            System.out.println("修改该条零件信息成功!"+"已自动更新到"+pn.getPathnameTwo()+"中");
        }
        db.closeAll();
        printPartEntry();


        FileInputStream fis=new FileInputStream(pn.getPathnameTwo());
        XSSFWorkbook workbook = new XSSFWorkbook();
        //输入编号打印零件名称和零件规格
        Workbook sheets2=new XSSFWorkbook(fis);
        Sheet sheet1 = sheets2.getSheetAt(0);
        XSSFRow row = null;
        XSSFCell cell = null;
        FileOutputStream fos=null;
        XSSFSheet sheet = workbook.createSheet("零件进库表");
        fos = new FileOutputStream(pn.getPathnameTwo());
        //设置标题
        row = sheet.createRow(0);
        cell = row.createCell(0);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 1);
        row.createCell(0).setCellValue("零件进库表信息");
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellStyle(style);
        sheet.addMergedRegion(region);
        for (int i = 0; i < two.getNumber().length; i++) {
            row = sheet.createRow(i+1);
            cell = row.createCell(i+1);
            if (i == 0) {
                row.createCell(0).setCellValue(two.getNumber()[0]);
                row.createCell(1).setCellValue("零件数量");
            } else {
                row.createCell(0).setCellValue(two.getNumber()[i]);
                if (number.equals(two.getNumber()[i])) {
                    two.amount[i]=amount;
                    two.setAmount(two.amount);
                    row.createCell(1).setCellValue(amount);
                }else{
                    row.createCell(1).setCellValue(two.getAmount()[i]);
                }
            }
        }
        workbook.write(fos);
        workbook.close();
        fos.close();
        fis.close();
        sheets2.close();

    }

    public static void deletePratEntry(ComponentsTwo two,Pathname pn) throws SQLException, IOException {
        String deleteSql="delete from partentry where number=?";;
        Scanner sc=new Scanner(System.in);
        System.out.println("输入删除的零件对应的零件编号:");
        String number = sc.next();
        DBUtil db=new DBUtil();
        db.getConnection();
        Object[] param=new Object[]{number};
        int c=db.executeUpdate(deleteSql,param);
        if (c!=0){
            System.out.println("删除该条零件信息成功!"+"已自动更新到"+pn.getPathnameTwo()+"中!");
        }
        db.closeAll();
        printPartEntry();


        int b=0;
        for (int i = 0; i < two.number.length; i++) {
            if(number.equals(two.getNumber()[i])){
                b=i;
            }
        }
        for (int i = b; i < two.number.length-1; i++) { //1 2
            two.number[i]=two.getNumber()[i+1];
            two.amount[i]=two.getAmount()[i+1];
            two.setNumber(two.number);
            two.setAmount(two.amount);
        }
        //将零件信息的数组长度缩小一个长度
        two.number=Arrays.copyOf(two.number,two.number.length-1);
        two.amount=Arrays.copyOf(two.amount,two.amount.length-1);

        //创建一个删除指定零件编号的零件进库表
        FileOutputStream fos=new FileOutputStream(new File(pn.getPathnameTwo()));
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("零件进库表");
        XSSFRow rowDelete=null;
        XSSFCell cellDelete=null;
        for (int i = 0; i < two.number.length+1; i++) {
            rowDelete=sheet.createRow(i);
            cellDelete=rowDelete.createCell(i);
            if(i==0){
                rowDelete = sheet.createRow(0);
                cellDelete = rowDelete.createCell(0);
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 1);
                rowDelete.createCell(0).setCellValue("零件进库表信息");
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cellDelete.setCellStyle(style);
                sheet.addMergedRegion(region);
            }else{
                rowDelete.createCell(0).setCellValue(two.getNumber()[i-1]);
                if(i==1){
                    rowDelete.createCell(1).setCellValue("零件数量");
                }else {
                    rowDelete.createCell(1).setCellValue(two.getAmount()[i-1]);
                }
            }
        }
        workbook.write(fos);
        workbook.close();
        fos.close();
    }

    public static void printPartEntry() throws SQLException {
        DBUtil db=new DBUtil();
        db.getConnection();
        String sql="select * from partentry";
        ResultSet rs=db.executeQuery(sql,null);
        System.out.println("零件编号  零件数量");
        while (rs.next()){
            System.out.println(rs.getString(1)+"       "+rs.getInt(2));
        }
        db.closeAll();
    }
}
