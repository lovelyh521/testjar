package excute;

import com.csvreader.CsvWriter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class CreateCSVFile {

    static Map<String,String> map = new HashMap<>();
    static int step = 0;
    static int totleCount = 0;

    public static void createCVS(String filePath, String fileName,String charSet,String date) throws IOException {
//        XSSFWorkbook workBook = new XSSFWorkbook();
//        XSSFSheet sheet = workBook.createSheet();

        // 创建CSV写对象
        CsvWriter csvWriter = null;


        /*csvWriter.close();*/




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
        boolean flag1 = false;
        boolean get = false;

        while (temp != null) {
           /* if(get){
                try{
                    String outFileName = map.get(Column.column7);
                    if(outFileName != null){
                        FileInputStream fs=new FileInputStream(".\\outputfile\\"+outFileName+".xlsx");

                    }
                }catch (Exception e){
                    System.out.println("文件不存在将生成");
                }finally {
                    System.out.println(step);
                    get = false;
                }
            }*/
            if(temp.startsWith("---------------------------------")){
                get = true;
                flag = true;
                String outFileName = crestFilePath(fileName, date);

                csvWriter = new CsvWriter(outFileName,',', Charset.forName("GB2312"));
                temp = br.readLine();
                continue;
            }
            if(temp.indexOf("No")!=-1){
                if(flag1){
                    System.out.println("生成"+i+"条数据");
                }
                flag1=true;
                totleCount += i;
                i=0;
                if(csvWriter != null){
                    csvWriter.close();
                }


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
            if(temp.startsWith("本次批处理")){
                csvWriter.close();
                csvWriter=null;
                break;
            }
            String[] content = new String[10];

            if(temp != null && !"".equals(temp.trim())){
                String[] split = temp.trim().split("\\s{2,}");
                String[] split1 = split[0].split(",");
                Column.column1 = split1[0].split("=")[0];
                Column.column2 = split1[1].split("=")[0];
                Column.column3 = split1[2].split("=")[0];
                Column.column4 = split1[3].split("=")[0];
                Column.column5 = split1[4].split("=")[0];

                if(flag){
                    if(step == 0){
                        // 写表头
                        String[] headers = {"",Column.column1,Column.column2,Column.column3,Column.column4,Column.column5,Column.column6,Column.column7};
                        csvWriter.writeRecord(headers);
                        flag = false;
                    }
                }

                content[0]=rXc0;
                for(int s1 = 0 ;s1<split1.length;s1++){
                    content[s1+1]=split1[s1].split("=")[1];
                }
                for (int j=1 ; j<split.length;j++){
                    String s = split[j];
                    content[j+5]=s;
                }

                csvWriter.writeRecord(content);
                i++;
            }
            temp = br.readLine();
        }
        System.out.println("当前文件共生成"+totleCount+"条数据");
    }
    static int j=0;


    private static String crestFilePath(String fileName, String date){
        String outFileName = "";
        date = "ss";
        if(date !=null){
//            SimpleDateFormat sf = new SimpleDateFormat(date);
//            String format = sf.format(new Date());
            String format = String.valueOf(System.currentTimeMillis());
            outFileName = fileName.substring(0, fileName.indexOf("."))+"_"+ Column.column7+"_"+format.substring(7);
        }else{
            outFileName = fileName.substring(0, fileName.indexOf("."))+"_"+ Column.column7;
        }

        File f = new File(".\\outputfile\\");
        if(!f.exists()){
            f.mkdir();
        }
        return ".\\outputfile\\"+(outFileName.replaceAll("\\\\","_").replaceAll("/","_")+".csv");
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


    public static void main(String[] args) {
        String a = "\\sdf".replaceAll("\\\\","_");
        System.out.println("a = " + a);
    }
}
