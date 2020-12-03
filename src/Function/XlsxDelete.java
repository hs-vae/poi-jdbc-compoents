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


import java.io.*;
import java.util.Arrays;
import java.util.Scanner;

/**
 *  删除功能:根据零件编号删除对应的零件(零件表,零件进库表,零件汇总表)信息
 */
public class XlsxDelete {

    //根据零件编号删除该零件表的零件信息的方法
    public  static void deleteXlsxOne(ComponentsOne one, Pathname pn) throws IOException {
        System.out.println("------零件表(删除前)------");
        XlsxOne.printXlsxOne(pn);
            Scanner sc=new Scanner(System.in);
            System.out.println("输入你想删除的零件信息对应的编号:");
            String deleteNumber=sc.next();
            int b=0;
        for (int i = 0; i < one.number.length; i++) {
            if(deleteNumber.equals(one.getNumber()[i])){
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
        //输出删除后的零件表信息
        System.out.println("------零件表(删除后)------");
        XlsxOne.printXlsxOne(pn);
    }

    //根据零件编号删除该零件进库表的零件信息的方法
    public static void deleteXlsxTwo(ComponentsTwo two, Pathname pn) throws IOException {
        System.out.println("------零件进库表(删除前)------");
        XlsxTwo.printXlsxTwo(pn);
        Scanner sc=new Scanner(System.in);
        System.out.println("输入你想删除的零件信息对应的零件编号:");
        String deleteNumber=sc.next();
        int b=0;
        for (int i = 0; i < two.number.length; i++) {
            if(deleteNumber.equals(two.getNumber()[i])){
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
        //输出删除后的零件进库表信息
        System.out.println("------零件进库表(删除后)------");
        XlsxTwo.printXlsxTwo(pn);
    }

    //根据零件编号删除该零件汇总表的零件信息的方法
    public static void deleteXlsxSum(ComponentsSum three, Pathname pn) throws IOException {
        System.out.println("------零件汇总表(删除前)------");
        XlsxSum.PrintSumXlsx(pn);
        Scanner sc=new Scanner(System.in);
        System.out.println("输入你想删除的零件信息对应的编号:");
        String deleteNumber=sc.next();
        int b=0;
        for (int i = 0; i < three.number.length; i++) {
            if(deleteNumber.equals(three.getNumber()[i])){
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
        //输出删除后的零件表信息
        System.out.println("------零件汇总表(删除后)------");
        XlsxSum.PrintSumXlsx(pn);
    }
}
