package com.qxn.test;

import java.io.FileOutputStream;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * @author  作者 :qxn
 * @date 创建时间：2018年1月25日 上午10:24:06
 * @version 1.0
 * @parameter
 * @since
 * @return 
 */
public class CreateExcel {

	public static void main(String[] args) throws Exception {
		
		HSSFWorkbook wb=new HSSFWorkbook();//创建一个工作簿
		HSSFSheet sheet=wb.createSheet("学生信息");//创建一个单子
		
		//创建单子样式
		HSSFCellStyle style=wb.createCellStyle();
		//添加具体样式
		style.setAlignment(style.ALIGN_CENTER);
		
		HSSFRow row=sheet.createRow(0);//创建第一行（表头）
		HSSFCell cell1=row.createCell(0);//创建第一列
		//为当前单元格添加样式
		cell1.setCellStyle(style);
		cell1.setCellValue("姓名");//为当前单元格赋值
		HSSFCell cell2=row.createCell(1);//创建第二列
		cell2.setCellStyle(style);
		cell2.setCellValue("性别");//为当前单元格赋值
		HSSFCell cell3=row.createCell(2);//创建第三列
		cell3.setCellStyle(style);
		cell3.setCellValue("年龄");//为当前单元格赋值
		HSSFCell cell4=row.createCell(3);//创建第四列
		cell4.setCellStyle(style);
		cell4.setCellValue("电话");//为当前单元格赋值
		
		List<StudentModel> list=new ArrayList<StudentModel>();
		for(int i=0;i<5;i++){
			StudentModel studentModel=new StudentModel();
			studentModel.setS_age(122+i);
			studentModel.setS_name("亚索"+i);
			studentModel.setS_phone("1380013800"+i);
			studentModel.setS_sex("男");
			list.add(studentModel);
		}
		
		for(int i=0;i<list.size();i++){
			StudentModel studentModel=list.get(i);
			HSSFRow row1=sheet.createRow(i+1);//创建一行
			HSSFCell cell11=row1.createCell(0);//创建一列
			cell11.setCellStyle(style);
			cell11.setCellValue(studentModel.getS_name());//为当前单元格赋值
			HSSFCell cell22=row1.createCell(1);
			cell22.setCellStyle(style);
			cell22.setCellValue(studentModel.getS_sex());
			HSSFCell cell33=row1.createCell(2);
			cell33.setCellStyle(style);
			cell33.setCellValue(studentModel.getS_age());
			HSSFCell cell44=row1.createCell(3);
			cell44.setCellStyle(style);
			cell44.setCellValue(studentModel.getS_phone());
		}
		FileOutputStream out=new FileOutputStream("d:/student.xls");//创建输出流，用于输出文件
		wb.write(out);//将文件写入具体路径
		out.close();
	}
}
