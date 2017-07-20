import Excute.TestFile;
import org.apache.poi.hssf.extractor.ExcelExtractor;

import java.io.*;

public class MyMain {
    public static void main(String[] args) {
//        String path = MyMain.class.getResource("/").getFile()+ "sourceFile";
        String path = System.getProperty("user.dir")+ "\\sourceFile";
        System.out.println("读取文件目录："+path);
//        System.out.println("读取文件目录：F:\\testjar\\out\\artifacts\\testjar_jar\\sourceFile");
        File file = new File(path);
        if(file != null){
            File[] files = file.listFiles();
            if(files != null){
                for(File fl:files){
                    if(fl.isFile()){
                        try {
                            TestFile.creat2007Excel(fl.getPath(),fl.getName());
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                }
            }else {
                System.out.println("读取文件错误:请创建 sourceFile 文件夹并把需要读取的文件放入其中(注意大小写)。");
            }

        }/*else{
            System.out.println("读取文件错误:请创建 sourceFile 文件夹并把需要读取的文件放入其中。");
        }*/
    }
}
