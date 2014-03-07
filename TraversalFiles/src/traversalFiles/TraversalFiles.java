package traversalFiles;

import java.io.File;
import java.io.FileFilter;

public class TraversalFiles {
	 public static void main(String[] args) {
	        File file = new File("E:\\这是测试text");
	        File[] files = file.listFiles(new FileFilter(){
	            public boolean accept(File pathname) {
	                // 判断文件名是目录 或 是word文档
	                if (pathname.getName().toString().endsWith(".doc")//判断文件为以下后缀或为文件夹
	                ||	pathname.getName().toString().endsWith(".DOC")
	                ||	pathname.getName().toString().endsWith(".DOC")
	                ||	pathname.getName().toString().endsWith(".DOCX")
	                ||	pathname.isDirectory()) {
	                    return true;
	                }
	                else
	                	return false;
	            }});
	            
	        for (File f : files) {
	            System.out.println(f.getName());
	        }
	        System.out.println("输出完毕");
	    }
}
