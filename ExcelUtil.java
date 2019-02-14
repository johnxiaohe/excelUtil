package com.cz.czUser.system.utils;

import org.apache.commons.fileupload.disk.DiskFileItem;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import java.io.*;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author hxy
 * Excel文件导入工具类
 */
public class ExcelUtil {
    private final static String excel2003L = ".xls";
    private final static String excel2007U = ".xlsx";

    /**
     * 获取实体对象返回属性名称
     */
    public java.lang.reflect.Field[] findEntityAllTypeName(Object obj) throws Exception {
        //反射获取类的字节码对象
        Class<? extends Object> cls = obj.getClass();
        return cls.getDeclaredFields();
    }

    /**
     * 下载文件
     */
    public String uploadFile(MultipartFile file) throws Exception {
        String filename = file.getOriginalFilename();
//        String filenamePrx = filename.split("\\.")[0];
        String path = "D:";
        String localPath = path + File.separator + "excel";
        File fileLocal = new File(localPath,filename);

        if (!fileLocal.exists() && !fileLocal.isDirectory()) {
            fileLocal.mkdirs();
        }
        file.transferTo(fileLocal);
//        InputStream in = file.getInputStream();
//        byte[] b = new byte[1024];
//        int length = in.read(b);
//
//        FileOutputStream outFile = new FileOutputStream(localPath + File.separator + filename);
//        outFile.write(b,0,length);
//
//        in.close();
//        outFile.close();

        return localPath + File.separator + filename;
    }

    /**
     * 根据文件选择excel版本
     */
    public Workbook chooseWorkbook(MultipartFile file) throws Exception {
        Workbook workbook = null;

        //吧MultipartFile转化为File
        File fo = new File(uploadFile(file));

        String filename = file.getOriginalFilename();
        String fileType = (filename.substring(filename.lastIndexOf("."), filename.length())).toLowerCase();
        if (excel2003L.equals(fileType)) {
            workbook = new HSSFWorkbook(FileUtils.openInputStream(fo));
        } else if (excel2007U.equals(fileType)) {
            workbook = new XSSFWorkbook(FileUtils.openInputStream(fo));
        } else {
            throw new Exception("解析文件格式错误");
        }
        return workbook;
    }
//}

    /**
     * @Author:和笑远
     * @Date：2019.1.29
     * @param file              上传文件输入流   表单上传MultipartFile方式
     * @param sheetname         sheetname   excel工作sheet表的名字   如果没有更改传sheet1
     * @param obj               要导入数据库对应的对象
     * @param attributeName     要导入数据库对应的字段名称  大小写要和pojo类对象属性一致  插入顺序要和excel表格顺序一致
     * @return
     * @throws Exception
     */
    public List<Object> importBaseExcel(MultipartFile file , String sheetname , Object obj,List<String> attributeName)throws Exception{
        Workbook workbook = null;
        try{
            //读取文件内容
            workbook = this.chooseWorkbook(file);
            //获取工作表
            Sheet sheet = workbook.getSheet(sheetname);

            //获取sheet中第一行行号
            int firstRowNum = sheet.getFirstRowNum();
            //获取sheet中最后一行行号
            int lastRowNum = sheet.getLastRowNum();

            //获取该实体类所有定义的属性 返回Field数组    该数组里
            java.lang.reflect.Field[] entityName = this.findEntityAllTypeName(obj);

            //获取类的名字
            String classname = obj.getClass().getName();
            //获取到该类的Class对象
            Class<?> clazz = Class.forName(classname);
            //将该类全部放到这个list中
            List<Object> list = new ArrayList<>();

            //循环插入数据    遍历excel表
            for(int i=firstRowNum+1;i<=lastRowNum;i++){
                Row row = sheet.getRow(i);

                //根据该类名生成java对象  生成该类的object对象
                Object pojo = clazz.newInstance();

                //除自增编号外，实体字段匹配sheet列   遍历每一列
                int attrM = 0;
                for(int j=0;j<entityName.length;j++){
                    String attributeNameMethod = entityName[j].getName();
                    if(attrM==attributeName.size()){

                        break;
                    }
                    //获取到的对象属性跟我们要导入的属性不一致就不装填 因为也没有这个值
                    if(!attributeName.contains(attributeNameMethod)){
                        continue;
                    }
                    //获取属性的名字,将属性的首字符大写，方便构造set方法 根据该类的属性名获取到该类对应属性的的方法
                    String name = "set"+entityName[j].getName().substring(0, 1).toUpperCase().concat(entityName[j].getName().substring(1).toLowerCase())+"";
                    //获取属性的类型
                    String type = entityName[j].getGenericType().toString();

                    Method m = null;
                    //getMethod只能调用public声明的方法，而getDeclaredMethod基本可以调用任何类型声明的方法
                    m = obj.getClass().getDeclaredMethod(name,entityName[j].getType());

                    Cell pname  = row.getCell(attrM);
                    //根据属性类型装入值
                    switch(type){
                        case "char":
                        case "java.lang.Character":
                        case "class java.lang.String":
                            //根据此方法将其值装填到对象的对应属性中
                            m.invoke(pojo,getVal(pname));
                            break;
                        case "int":
                        case "class java.lang.Integer":
                            //根据此方法将其值装填到对象的对应属性中
                            m.invoke(pojo,Integer.valueOf(getVal(pname)));
                            break;
                        case "class java.util.Date":
                            //根据此方法将其值装填到对象的对应属性中
                            m.invoke(pojo,getVal(pname));
                            break;
                        case "float":
                        case "double":
                        case "java.lang.Double":
                        case "java.lang.Float":
                        case "java.lang.Long":
                        case "java.lang.Short":
                        case "java.math.BigDecimal":
                            //根据此方法将其值装填到对象的对应属性中
                            m.invoke(pojo,Double.valueOf(getVal(pname)));
                            break;
                        default:
                            break;
                    }
                    attrM++;//最后j++装填下一列的值
                }
                list.add(pojo);
            }
            return list;
        }catch (Exception e){
            e.printStackTrace();
        }
        return null;
    }


    /**
     * 处理类型
     * @param cell
     * @return
     */
    public static String getVal(Cell cell) {
        if (null != cell) {

            switch (cell.getCellType()) {
                case XSSFCell.CELL_TYPE_NUMERIC: // 数字

                    String val = cell.getNumericCellValue()+"";
                    int index = val.indexOf(".");

                    if(Integer.valueOf(val.substring(index+1)) == 0){
                        DecimalFormat df = new DecimalFormat("0");//处理科学计数法
                        return df.format(cell.getNumericCellValue());
                    }
                    return cell.getNumericCellValue()+"";//double
                case XSSFCell.CELL_TYPE_STRING: // 字符串
                    return cell.getStringCellValue() + "";
                case XSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                    return cell.getBooleanCellValue() + "";
                case XSSFCell.CELL_TYPE_FORMULA: // 公式

                    try{
                        if(HSSFDateUtil.isCellDateFormatted(cell)){
                            Date date = cell.getDateCellValue();
                            return (date.getYear() + 1900) + "-" + (date.getMonth() + 1) +"-" + date.getDate();
                        }else{
                            return String.valueOf((int)cell.getNumericCellValue());
                        }
                    }catch (IllegalStateException e) {
                        return  String.valueOf(cell.getRichStringCellValue());
                    }
                case XSSFCell.CELL_TYPE_BLANK: // 空值
                    return "";
                case XSSFCell.CELL_TYPE_ERROR: // 故障
                    return "";
                default:
                    return "未知类型   ";
            }
        } else {
            return "";
        }
    }

}
