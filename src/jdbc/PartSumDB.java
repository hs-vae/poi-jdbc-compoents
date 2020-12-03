package jdbc;


import Entity.ComponentsSum;
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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.Scanner;

class G implements Comparable<G>{ //实现comparable接口,复写comparato方法
    String number;
    String name;
    String size;
    int amount;
    @Override
    public int compareTo(G g) {
        if (this.number.compareTo(g.number) <0 )
            return 1;
        else
            return -1;
    }
}
/**
 * 操作零件汇总表在数据库中对应的数据类
 */
public class PartSumDB {


    public static void insertAllPartSum(ComponentsSum three) throws SQLException {
        String insertSql="insert into partsum(number,name,size,amount)values(?,?,?,?)";
        DBUtil db=new DBUtil();
        db.getConnection();
        int a=0;
        for (int i = 1; i < three.number.length; i++) {
            Object[] param =new Object[]{three.getNumber()[i],three.getName()[i],three.getSize()[i],three.getAmount()[i]};
            a=db.executeUpdate(insertSql,param);
        }
        if(a!=0){
            System.out.println("零件汇总表.xlsx信息导入数据库partsum表成功!");
        }
        db.closeAll();
        printPartSum();
    }

    public static void insertPartSum(ComponentsSum three, Pathname pn) throws IOException, SQLException {
        String insertSql = "insert into partsum(number,name,size,amount)values(?,?,?,?)";
        DBUtil db = new DBUtil();
        db.getConnection();
        Scanner sc = new Scanner(System.in);
        System.out.println("请输入要插入的零件编号:");
        String number = sc.next();
        System.out.println("请输入要插入的零件名称:");
        String name = sc.next();
        System.out.println("请输入要插入的零件规格:");
        String size = sc.next();
        System.out.println("请输入要插入的零件数量:");
        int amount =sc.nextInt();


        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("零件汇总表");
        FileOutputStream fos=new FileOutputStream(new File(pn.getPathnameSum()));
        Workbook sheets=null;
        XSSFRow row=null;
        XSSFCell cell=null;
        //增加零件表每个零件信息数组的长度+1
        three.number = Arrays.copyOf(three.number,three.number.length+1);
        three.name = Arrays.copyOf(three.name,three.name.length+1);
        three.size = Arrays.copyOf(three.size,three.size.length+1);
        three.amount = Arrays.copyOf(three.amount,three.amount.length+1);
        //将汇总表的零件信息全部存储在数组g里面
        G g [] = new G[three.number.length-1];
        for (int i = 0; i < g.length; i++) {   //g.length==6
            g[i] = new G();
            g[i].number= three.getNumber()[i+1];
            g[i].name = three.getName()[i+1];
            g[i].size= three.getSize()[i+1];
            g[i].amount = three.getAmount()[i+1];
            if(i==(g.length-1)){
                g[i].number = number;
                g[i].name = name;
                g[i].size = size;
                g[i].amount = amount;
            }
        }
        //从大到小对编号进行排序
        Arrays.sort(g);
        //想得到从小到大的编号,所以倒置循环一下,将零件信息又存储在three对象的成员变量数组中
        int e=1;
        for (int i = g.length; i >= 1; i--) {
            three.number[e]=g[i-1].number;
            three.name[e]=g[i-1].name;
            three.size[e]=g[i-1].size;
            three.amount[e]=g[i-1].amount;
            ++e;
        }
        for (int i = 0; i < three.number.length+1; i++) {
            row = sheet.createRow(i);
            cell = row.createCell(i);
            if (i == 0) {
                //设置单元格范围地址,参数分别表示为(起始行号,终止行号,起始列号,终止列号)
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 3);
                row.createCell(0).setCellValue("零件汇总表信息");
                //设置水平居中
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cell.setCellStyle(style);
                sheet.addMergedRegion(region);
                three.number[0]="零件编号";
                three.name[0]="零件名称";
                three.size[0]="零件规格";
                three.amount[0]=0;
            }  else {
                three.setNumber(three.number);
                three.setName(three.name);
                three.setSize(three.size);
                three.setAmount(three.amount);
                //在数组尾部添加新增的零件信息
                row.createCell(0).setCellValue(three.getNumber()[i-1]);
                row.createCell(1).setCellValue(three.getName()[i-1]);
                row.createCell(2).setCellValue(three.getSize()[i-1]);
                if(i==1){
                    row.createCell(3).setCellValue("零件数量");
                }else{
                    row.createCell(3).setCellValue(three.getAmount()[i-1]);
                }
            }
        }
        System.out.println("添加成功,已存入" + pn.getPathnameSum() + "目录中!");
        workbook.write(fos);
        workbook.close();
        fos.close();

        String sql="delete from partsum";
        db.executeUpdate(sql,null);
        int a=0;
        for (int i = 1; i < three.number.length; i++) {
            Object[] param =new Object[]{three.getNumber()[i],three.getName()[i],three.getSize()[i],three.getAmount()[i]};
            a=db.executeUpdate(insertSql,param);
        }
        if (a!=0){
            System.out.println("插入该条数据成功!按编号自动排序+"+",已自动更新到"+pn.getPathnameSum()+"中");
        }
        db.closeAll();
        printPartSum();

    }


    public static void selectPartSum() throws SQLException {
        String selectSql="select number,name,size,amount from partsum where number=?";
        DBUtil db=new DBUtil();
        db.getConnection();
        Scanner sc=new Scanner(System.in);
        System.out.println("输入你想查询的零件编号:");
        String number=sc.next();
        Object[] param=new Object[]{number};
        ResultSet rs=db.executeQuery(selectSql,param);
        while (rs.next()){
            System.out.println(rs.getString(1)+"       "+rs.getString(2)+"       "+rs.getString(3)+"       "+rs.getInt(4));
        }
        db.closeAll();

    }
    public static void updatePartSum(ComponentsSum three,Pathname pn) throws SQLException, IOException {
        String updateSql="update partsum set name=?,size=?,amount=? where number=?";
        Scanner sc=new Scanner(System.in);
        System.out.println("输入你想修改零件对应的编号:");
        String number=sc.next();
        System.out.println("输入修改后的零件名称:");
        String name=sc.next();
        System.out.println("输入修改后的零件规格:");
        String size=sc.next();
        System.out.println("输入修改后的零件数量:");
        int amount = sc.nextInt();

        DBUtil db=new DBUtil();
        db.getConnection();
        Object[] param=new Object[]{name,size,amount,number};
        int b=db.executeUpdate(updateSql,param);
        if (b!=0){
            System.out.println("从partsum表修改该条零件信息成功!"+"已自动更新到"+pn.getPathnameSum()+"中!");
        }
        db.closeAll();
        printPartSum();


        FileInputStream fis=new FileInputStream(pn.getPathnameSum());
        XSSFWorkbook workbook = new XSSFWorkbook();
        Workbook sheets2=new XSSFWorkbook(fis);
        Sheet sheet1 = sheets2.getSheetAt(0);
        XSSFRow row = null;
        XSSFCell cell = null;
        FileOutputStream fos=null;
        XSSFSheet sheet = workbook.createSheet("零件汇总表");
        fos = new FileOutputStream(pn.getPathnameSum());
        //设置标题
        row = sheet.createRow(0);
        cell = row.createCell(0);
        CellRangeAddress region = new CellRangeAddress(0, 0, 0, 3);
        row.createCell(0).setCellValue("零件汇总表信息");
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellStyle(style);
        sheet.addMergedRegion(region);
        //设置表头和表体
        for (int i = 0; i < three.getNumber().length; i++) {
            row = sheet.createRow(i+1);
            cell = row.createCell(i+1);
            if (i == 0) {
                row.createCell(0).setCellValue(three.getNumber()[0]);
                row.createCell(1).setCellValue(three.getName()[0]);
                row.createCell(2).setCellValue(three.getSize()[0]);
                row.createCell(3).setCellValue("零件数量");
            } else {
                row.createCell(0).setCellValue(three.getNumber()[i]);
                if (number.equals(three.getNumber()[i])) {
                    three.name[i]=name;
                    three.size[i]=size;
                    three.amount[i]=amount;
                    three.setName(three.name);
                    three.setSize(three.size);
                    three.setAmount(three.amount);
                    row.createCell(1).setCellValue(name);
                    row.createCell(2).setCellValue(size);
                    row.createCell(3).setCellValue(amount);
                }else{
                    row.createCell(1).setCellValue(three.getName()[i]);
                    row.createCell(2).setCellValue(three.getSize()[i]);
                    row.createCell(3).setCellValue(three.getAmount()[i]);
                }
            }
        }
        workbook.write(fos);
        workbook.close();
        fos.close();
        fis.close();
        sheets2.close();

    }
    public static void deletePratSum(ComponentsSum three,Pathname pn) throws SQLException, IOException {
        String deleteSql="delete from partsum where number=?";
        Scanner sc=new Scanner(System.in);
        System.out.println("输入删除的零件对应的零件编号:");
        String number = sc.next();
        DBUtil db=new DBUtil();
        db.getConnection();
        Object[] param=new Object[]{number};
        int c=db.executeUpdate(deleteSql,param);
        if (c!=0){
            System.out.println("从partsum表删除该条零件信息成功!"+"已自动更新到"+pn.getPathnameSum()+"中!");
        }
        db.closeAll();
        printPartSum();

        int b=0;
        for (int i = 0; i < three.number.length; i++) {
            if(number.equals(three.getNumber()[i])){
                b=i;
            }
        }
        for (int i = b; i < three.number.length-1; i++) { //1 2
            three.number[i]=three.getNumber()[i+1];
            three.name[i]=three.getName()[i+1];
            three.size[i]=three.getSize()[i+1];
            three.amount[i]=three.getAmount()[i+1];
            three.setNumber(three.number);
            three.setName(three.name);
            three.setSize(three.size);
            three.setAmount(three.amount);
        }
        //将零件信息的数组长度缩小一个长度
        three.number=Arrays.copyOf(three.number,three.number.length-1);
        three.name=Arrays.copyOf(three.name,three.name.length-1);
        three.size=Arrays.copyOf(three.size,three.size.length-1);
        three.amount=Arrays.copyOf(three.amount,three.amount.length-1);
        //创建一个删除指定零件编号的零件表
        FileOutputStream fos=new FileOutputStream(new File(pn.getPathnameSum()));
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("零件汇总表");
        XSSFRow rowDelete=null;
        XSSFCell cellDelete=null;
        for (int i = 0; i < three.number.length+1; i++) {
            rowDelete=sheet.createRow(i);
            cellDelete=rowDelete.createCell(i);
            if(i==0){
                rowDelete = sheet.createRow(0);
                cellDelete = rowDelete.createCell(0);
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, 3);
                rowDelete.createCell(0).setCellValue("零件汇总表信息");
                CellStyle style = workbook.createCellStyle();
                style.setAlignment(HorizontalAlignment.CENTER);
                cellDelete.setCellStyle(style);
                sheet.addMergedRegion(region);
            }else{
                rowDelete.createCell(0).setCellValue(three.getNumber()[i-1]);
                rowDelete.createCell(1).setCellValue(three.getName()[i-1]);
                rowDelete.createCell(2).setCellValue(three.getSize()[i-1]);
                if(i==1){
                    rowDelete.createCell(3).setCellValue("零件数量");
                }else {
                    rowDelete.createCell(3).setCellValue(three.getAmount()[i-1]);
                }
            }
        }
        workbook.write(fos);
        workbook.close();
        fos.close();

    }
    public static void printPartSum() throws SQLException {
        DBUtil db=new DBUtil();
        db.getConnection();
        String sql="select * from partsum";
        ResultSet rs=db.executeQuery(sql,null);
        System.out.println("零件编号  零件名称  零件规格  零件数量");
        while (rs.next()){
            System.out.println(rs.getString(1)+"       "+rs.getString(2)+"       "+rs.getString(3)+"       "+rs.getInt(4));
        }
        db.closeAll();
    }
}
