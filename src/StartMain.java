import excute.CreateFile;

import java.io.*;

public class StartMain {
    public static void main(String[] args) {
        String chatset = "GB2312";
        if(args.length>0){
            chatset = args[0];
        }

//        String path = "d:\\testjar\\out\\artifacts\\testjar_jar\\sourceFile";
        String path = System.getProperty("user.dir")+ "\\sourceFile";
        System.out.println("读取文件目录："+path);
        File file = new File(path);
        if(file != null){
            File[] files = file.listFiles();
            if(files != null){
                for(File fl:files){
                    if(fl.isFile()){
                        try {
                            CreateFile.creat2007Excel(fl.getPath(),fl.getName(),chatset);
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
