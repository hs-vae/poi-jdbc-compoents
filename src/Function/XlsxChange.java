package Function;

import Entity.ComponentsOne;
import Entity.ComponentsSum;
import Entity.ComponentsTwo;
import Entity.Pathname;
import Generate.XlsxOne;
import Generate.XlsxSum;
import Generate.XlsxTwo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
/**
 *  修改功能类:
 *  零件表根据零件编号修改对应的零件名称和零件规格
 *  零件进库表根据零件编号修改对应的零件数量
 *  零件汇总表根据零件编号修改对应的零件名称,零件规格,零件数量
 */

public class XlsxChange {



    //修改零件表的某个信息方法
    public  static void changeXlsxOne(ComponentsOne one, Pathname pn) throws IOException {
        System.out.println("------零件表(修改前)------");
        XlsxOne.printXlsxOne(pn);
        FileInputStream fis=new FileInputStream(pn.getPathnameOne());
        XSSFWorkbook workbook = new XSSFWorkbook();
        //输入编号打印零件名称和零件规格
        Workbook sheets1=new XSSFWorkbook(fis);
        Sheet sheet1 = sheets1.getSheetAt(0);
        Scanner sc=new Scanner(System.in);
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
        //设置表头和表体
        System.out.println("输入想要修改的零件对应的编号:");
        String idXlsx=sc.next();
        System.out.println("输入修改后的零件名字:");
        String nameChange=sc.next();
        System.out.println("输入修改后的零件规格:");
        String sizeChange=sc.next();
        for (int i = 0; i < one.getNumber().length; i++) {
            row = sheet.createRow(i + 1);
            cell = row.createCell(i + 1);
            if (i == 0) {
                row.createCell(0).setCellValue(one.getNumber()[0]);
                row.createCell(1).setCellValue(one.getName()[0]);
                row.createCell(2).setCellValue(one.getSize()[0]);
            } else {
                row.createCell(0).setCellValue(one.getNumber()[i]);
                if (idXlsx.equals(one.getNumber()[i])) {
                    one.name[i] = nameChange;
                    one.size[i] = sizeChange;
                    one.setName(one.name);
                    one.setSize(one.size);
                    row.createCell(1).setCellValue(nameChange);
                    row.createCell(2).setCellValue(sizeChange);
                } else {
                    row.createCell(1).setCellValue(one.getName()[i]);
                    row.createCell(2).setCellValue(one.getSize()[i]);
                }
            }
        }
        workbook.write(fos);
        workbook.close();
        fos.close();
        System.out.println("修改成功,已存入" +pn.getPathnameOne()+ "目录中!");
        //打印输出修改后的零件表信息
        System.out.println("------零件表(修改后)------");
        XlsxOne.printXlsxOne(pn);
        fis.close();
        sheets1.close();
    }



    //修改零件进库表中某个零件信息的方法
    public static void changeXlsxTwo(ComponentsTwo two, Pathname pn) throws IOException {
        System.out.println("------零件进库表(添加前)------");
        XlsxTwo.printXlsxTwo(pn);
        FileInputStream fis=new FileInputStream(pn.getPathnameTwo());
        XSSFWorkbook workbook = new XSSFWorkbook();
        //输入编号打印零件名称和零件规格
        Workbook sheets2=new XSSFWorkbook(fis);
        Sheet sheet1 = sheets2.getSheetAt(0);
        Scanner sc=new Scanner(System.in);
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
        //设置表头和表体
        System.out.println("输入想要修改的零件对应的编号:");
        String idXlsx=sc.next();
        System.out.println("输入修改后的零件数量:");
        int amountChange=sc.nextInt();
            for (int i = 0; i < two.getNumber().length; i++) {
                row = sheet.createRow(i+1);
                cell = row.createCell(i+1);
                if (i == 0) {
                    row.createCell(0).setCellValue(two.getNumber()[0]);
                    row.createCell(1).setCellValue("零件数量");
                } else {
                    row.createCell(0).setCellValue(two.getNumber()[i]);
                    if (idXlsx.equals(two.getNumber()[i])) {
                        two.amount[i]=amountChange;
                        two.setAmount(two.amount);
                        row.createCell(1).setCellValue(amountChange);
                    }else{
                        row.createCell(1).setCellValue(two.getAmount()[i]);
                    }
                }
            }
        workbook.write(fos);
        workbook.close();
        fos.close();
        System.out.println("修改成功,已存入" +pn.getPathnameTwo()+ "目录中!");
        //打印输出零件进库表信息
        System.out.println("------零件进库表(添加前)------");
        XlsxTwo.printXlsxTwo(pn);
        fis.close();
        sheets2.close();
    }

    //修改零件汇总表零件信息的方法
    public static void changeXlsxSum(ComponentsSum three, Pathname pn) throws IOException {
        System.out.println("------零件汇总表(添加前)------");
        XlsxSum.PrintSumXlsx(pn);
        FileInputStream fis=new FileInputStream(pn.getPathnameSum());
        XSSFWorkbook workbook = new XSSFWorkbook();
        Workbook sheets2=new XSSFWorkbook(fis);
        Sheet sheet1 = sheets2.getSheetAt(0);
        Scanner sc=new Scanner(System.in);
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
        System.out.println("输入想要修改的零件对应的编号:");
        String idXlsx=sc.next();
        //输入需要修改的零件名称,规格,数量
        System.out.println("输入修改后的零件名称");
        String nameChange=sc.next();
        System.out.println("输入修改后的零件规格");
        String sizeChange=sc.next();
        System.out.println("输入修改后的零件数量");
        int amountChange= sc.nextInt();
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
                    if (idXlsx.equals(three.getNumber()[i])) {
                        three.name[i]=nameChange;
                        three.size[i]=sizeChange;
                        three.amount[i]=amountChange;
                        three.setName(three.name);
                        three.setSize(three.size);
                        three.setAmount(three.amount);
                        row.createCell(1).setCellValue(nameChange);
                        row.createCell(2).setCellValue(sizeChange);
                        row.createCell(3).setCellValue(amountChange);
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
        System.out.println("修改成功,已存入" +pn.getPathnameSum()+ "目录中!");
        //打印修改后的零件汇总表信息
        System.out.println("------零件汇总表(添加后且自动排序)------");
        XlsxSum.PrintSumXlsx(pn);
        fis.close();
        sheets2.close();
    }
}

