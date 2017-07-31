import excute.CreateCSVFile;

import java.io.*;

public class StartMain {
    public static void main(String[] args) {
        String charset = "GB2312";
        String date = "";
        /*if(args.length>0){
            chatset = args[0];
        }*/
        if(args.length > 0){
            for (int i = 0;i < args.length;i++) {
                if("date".equals(args[i])){
                    date = args[i+1];
                }
                if("charset".equals(args[i])){
                    charset = args[i+1];
                }
            }
        }


//        String path = "f:\\testjar\\out\\artifacts\\testjar_jar\\sourceFile";
        String path = System.getProperty("user.dir")+ "\\sourceFile";
        System.out.println("读取文件目录："+path);
        File file = new File(path);
        if(file != null){
            File[] files = file.listFiles();
            if(files != null){
                for(File fl:files){
                    if(fl.isFile()){
                        try {
//                            CreateExcelFile.creat2007Excel(fl.getPath(),fl.getName(),charset,date);
                            CreateCSVFile.createCVS(fl.getPath(),fl.getName(),charset,date);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                }
                System.out.println("全部文件导出成功");
            }else {
                System.out.println("读取文件错误:请创建 sourceFile 文件夹并把需要读取的文件放入其中(注意大小写)。");
            }

        }/*else{
            System.out.println("读取文件错误:请创建 sourceFile 文件夹并把需要读取的文件放入其中。");
        }*/
    }
}
