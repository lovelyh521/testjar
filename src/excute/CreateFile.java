package excute;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class CreateFile {
    public static String readFromFile(String filePath) throws IOException {
        File file = new File(filePath);
        if (!file.exists() || file.isDirectory())
            throw new FileNotFoundException();
        FileInputStream fin = new FileInputStream(file);
        InputStreamReader isr = new InputStreamReader(fin, "GB2312");
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
    static Map<String,String> map = new HashMap<>();
    public static void creat2007Excel(String filePath, String fileName,String charSet,String date) throws IOException {
        XSSFWorkbook workBook = new XSSFWorkbook();
        XSSFSheet sheet = workBook.createSheet();


        File file = new File(filePath);
        if (!file.exists() || file.isDirectory()){
            throw new FileNotFoundException();
        }

        FileInputStream fin = new FileInputStream(file);
        InputStreamReader isr = new InputStreamReader(fin, charSet);
        BufferedReader br = new BufferedReader(isr);
        String temp = null;
        temp = br.readLine();
        String rXc0 = "";

        int i = 0;
        boolean flag = false;
        boolean get = false;
        int step = 0;

        while (temp != null) {
            if(get){
                try{
                    String outFileName = map.get(fileName);
                    if(outFileName != null){
                        FileInputStream fs=new FileInputStream("\\outputfile\\"+outFileName+".xlsx");
                        XSSFWorkbook wb=new XSSFWorkbook(fs);
                        XSSFSheet sheet1=wb.getSheetAt(0);  //获取到工作表，因为一个excel可能有多个工作表
                        step = sheet1.getLastRowNum();
                    }
                }catch (Exception e){
                    e.printStackTrace();
                }finally {
                    get = false;
                }
            }
            if(temp.startsWith("--------------------")){
                get = true;
            }
            if(temp.indexOf("No")!=-1){
                if(flag){
                    excute(fileName,workBook,sheet,i,date);
                    i = 0;
                }

                flag = true;
                temp = br.readLine();
                continue;
            }

            if(temp.indexOf("MO SDR_OMMB") != -1){
                rXc0 = temp.replace("-","");
                temp = br.readLine();
                continue;
            }
            if( temp.trim().equals("结果")|| temp.indexOf("管理对象标识")!=-1 || temp.indexOf("-----")!=-1){
                if(temp.indexOf("管理对象标识")!=-1){
                    String[] split = temp.split("\\s+");
                    Column.column6 = split[1];
                    Column.column7 = split[2];
                    if(split.length>3){
                        Column.column7 += " "+split[3];
                    }
                }
                temp = br.readLine();
                continue;
            }
            if(temp.indexOf("本次批处理") !=-1){
                excute(fileName,workBook,sheet,i,date);
                break;
            }
            XSSFRow row = sheet.createRow(i+1+step);// 创建一个行对象
            if(temp != null && !"".equals(temp.trim())){
                String[] split = temp.trim().split("\\s+");
                String[] split1 = split[0].split(",");
                XSSFCell cell1 = row.createCell(0);// 创建单元格
                cell1.setCellValue(rXc0);
                Column.column1 = split1[0].split("=")[0];
                Column.column2 = split1[1].split("=")[0];
                Column.column3 = split1[2].split("=")[0];
                Column.column4 = split1[3].split("=")[0];
                Column.column5 = split1[4].split("=")[0];



                for(int s1 = 0 ;s1<split1.length;s1++){
                    XSSFCell cell2 = row.createCell(s1+1);// 创建单元格
                    cell2.setCellValue(split1[s1].split("=")[1]);
                }
                for (int j=1 ; j<split.length;j++){
                    XSSFCell cell = row.createCell(5+j);// 创建单元格
                    cell.setCellValue(split[j]);
                }

                i++;
            }

            temp = br.readLine();
        }

    }
    private static void excute( String fileName, XSSFWorkbook workBook, XSSFSheet sheet, int i,String date) throws IOException {

        XSSFRow rowTitle = sheet.createRow(0);// 创建一个行对象(表头)
        rowTitle.createCell(1).setCellValue(Column.column1);
        rowTitle.createCell(2).setCellValue(Column.column2);
        rowTitle.createCell(3).setCellValue(Column.column3);
        rowTitle.createCell(4).setCellValue(Column.column4);
        rowTitle.createCell(5).setCellValue(Column.column5);
        rowTitle.createCell(6).setCellValue(Column.column6);
        rowTitle.createCell(7).setCellValue(Column.column7);


        System.out.println("共转换 "+i+" 条数据");
        // 文件输出流

        String outFileName =crestFileName(fileName,date);
        File f = new File(".\\outputfile\\");
        if(!f.exists()){
            f.mkdir();
        }
        for(int k=0 ;k<8;k++){
            sheet.autoSizeColumn(k);
        }
        FileOutputStream os = new FileOutputStream(".\\outputfile\\"+outFileName+".xlsx");
//        FileOutputStream os = new FileOutputStream("d:\\"+outFileName+".xlsx");
        workBook.write(os);// 将文档对象写入文件输出流

        os.close();// 关闭文件输出流
        System.out.println("创建成功:.\\outputfile\\"+outFileName+".xlsx");
        map.put(fileName,outFileName);
        workBook = new XSSFWorkbook();
        sheet = workBook.createSheet();// 创建一个工作薄对象
    }


    private static String crestFileName(String fileName,String date){
        String outFileName = "";
        if(date !=null){
            SimpleDateFormat sf = new SimpleDateFormat(date);
            String format = sf.format(new Date());
            outFileName = fileName.substring(0, fileName.indexOf("."))+"-"+Column.column5+format;
        }else{
            outFileName = fileName.substring(0, fileName.indexOf("."))+"-"+Column.column5;
        }
        return outFileName;
    }

    private static class Column {
        static String column1 = "";
        static String column2 = "";
        static String column3 = "";
        static String column4 = "";
        static String column5 = "";
        static String column6 = "";
        static String column7 = "";
    }
}
