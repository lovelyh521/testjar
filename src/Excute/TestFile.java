package Excute;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.Arrays;
import java.util.Date;

public class TestFile {
    public static String readFromFile(String filePath) throws IOException {
        File file = new File(filePath);
        if (!file.exists() || file.isDirectory())
            throw new FileNotFoundException();
        FileInputStream fin = new FileInputStream(file);
        InputStreamReader isr = new InputStreamReader(fin, "UTF-8");
        BufferedReader br = new BufferedReader(isr);
        String temp = null;
        StringBuffer sb = new StringBuffer();
        temp = br.readLine();
        while (temp != null) {
            sb.append(temp + "\r\n");
            temp = br.readLine();
        }
        return sb.toString();
    }


    /**
     * 创建2007版Excel文件
     *
     * @throws FileNotFoundException
     * @throws IOException
     */
    public static void creat2007Excel(String filePath, String fileName) throws IOException {
        XSSFWorkbook workBook = new XSSFWorkbook();
        XSSFSheet sheet = workBook.createSheet();// 创建一个工作薄对象


        File file = new File(filePath);
        if (!file.exists() || file.isDirectory()){
            throw new FileNotFoundException();
        }

        FileInputStream fin = new FileInputStream(file);
        InputStreamReader isr = new InputStreamReader(fin, "UTF-8");
        BufferedReader br = new BufferedReader(isr);
        String temp = null;
//        StringBuffer sb = new StringBuffer();
        temp = br.readLine();
        int i = 0;
        while (temp != null) {

            XSSFRow row = sheet.createRow(i);// 创建一个行对象
            if(temp != null){
                String[] split = temp.trim().split("\\s+");
                for (int j=0 ; j<split.length;j++){
                    XSSFCell cell = row.createCell(j);// 创建单元格
                    cell.setCellValue(split[j]);
                }
                i++;
            }
            temp = br.readLine();
        }

        // 文件输出流
        String outFileName = fileName.substring(0, fileName.indexOf("."));

        File f = new File(".\\outputfile\\");
        if(!f.exists()){
            f.mkdir();
        }
        FileOutputStream os = new FileOutputStream(".\\outputfile\\"+outFileName+".xlsx");
        workBook.write(os);// 将文档对象写入文件输出流

        os.close();// 关闭文件输出流
        System.out.println("创建成功:.\\outputfile\\"+outFileName+".xlsx");
    }


}
