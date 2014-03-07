package allWordToTxt;

import java.io.File;
import java.io.FileFilter;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.*; 

public class AllWordToTxt {
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
	        
	        
	        for (File theseFiles : files){
	        	 //指定被转换文件的完整路径。 我这里的意图是把pdf转为txt  
	            String path = new String("E:\\这是测试text\\"+theseFiles.getName().toString());  
	            //根据路径创建文件对象  
	            File docFile=new File(path);  
	            //获取文件名（包含扩展名）  
	            String filename=docFile.getName();  
	            //过滤掉文件名中的扩展名  
	            //int filenamelength = filename.length();  
	            int dotposition=filename.indexOf(".");  
	            filename=filename.substring(0,dotposition);  
	              
	            //设置输出路径，一定要包含输出文件名（不含输出文件的扩展名）  
	            String savepath = new String ("E:\\这是测试text\\"+filename);    
	              
	            //启动Word程序  
	            ActiveXComponent app = new ActiveXComponent("Word.Application");          
	            //接收输入文件和输出文件的路径  
	            String inFile = path;  
	            String tpFile = savepath;  
	            //设置word不可见  
	            app.setProperty("Visible", new Variant(false));  
	            //这句不懂  
	            Object docs = app.getProperty("Documents").toDispatch();  
	            //打开输入的doc文档  
	            Object doc = Dispatch.invoke((Dispatch) docs,"Open", Dispatch.Method, new Object[]{inFile,new Variant(false), new Variant(true)}, new int[1]).toDispatch();  
	              
	            //另存文件, 其中Variant(n)参数指定另存为的文件类型，详见代码结束后的文字  
	            Dispatch.invoke((Dispatch) doc,"SaveAs", Dispatch.Method, new Object[]{tpFile,new Variant(2)}, new int[1]);  
	            //这句也不懂  
	            Variant f = new Variant(false);  
	            //关闭并退出  
	            Dispatch.call((Dispatch) doc, "Close", f);  
	            app.invoke("Quit", new Variant[] {});  
	            System.out.println("转换完毕。");  
	        }
	    }
}
