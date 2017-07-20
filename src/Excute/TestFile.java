package Excute;

import org.apache.poi.xssf.usermodel.*;

import java.io.*;

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
        temp = br.readLine();
        System.out.println("temp1 = " + temp);
        String rXc0 = "";

        XSSFRow rowTitle = sheet.createRow(0);// 创建一个行对象(表头)


        int i = 0;
        while (temp != null) {
            if(temp.indexOf("MO SDR_OMMB") != -1){
                rXc0 = temp.replace("-","");
                System.out.println("rXc0:"+rXc0);
                continue;
            }
            if(temp.startsWith("No") || temp.startsWith("结果") || temp.startsWith("管理对象标识") || temp.startsWith("-----")){

                System.out.println("true = " + true+"sdf"+i);
                System.out.println("temp = " + temp);
                continue;
            }
            XSSFRow row = sheet.createRow(i+1);// 创建一个行对象
            if(temp != null){
                System.out.println("temp = " + temp);
                String[] split = temp.trim().split("\\s+");
                String[] split1 = split[0].split(",");
                XSSFCell cell1 = row.createCell(0);// 创建单元格
                cell1.setCellValue(rXc0);
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
