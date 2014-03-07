package allWordToTxt;

import java.io.File;
import java.io.FileFilter;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.*; 

public class AllWordToTxt {
	 public static void main(String[] args) {
	        File file = new File("E:\\���ǲ���text");
	        File[] files = file.listFiles(new FileFilter(){
	            public boolean accept(File pathname) {
	                // �ж��ļ�����Ŀ¼ �� ��word�ĵ�
	                if (pathname.getName().toString().endsWith(".doc")//�ж��ļ�Ϊ���º�׺��Ϊ�ļ���
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
	        System.out.println("������");
	        
	        
	        for (File theseFiles : files){
	        	 //ָ����ת���ļ�������·���� ���������ͼ�ǰ�pdfתΪtxt  
	            String path = new String("E:\\���ǲ���text\\"+theseFiles.getName().toString());  
	            //����·�������ļ�����  
	            File docFile=new File(path);  
	            //��ȡ�ļ�����������չ����  
	            String filename=docFile.getName();  
	            //���˵��ļ����е���չ��  
	            //int filenamelength = filename.length();  
	            int dotposition=filename.indexOf(".");  
	            filename=filename.substring(0,dotposition);  
	              
	            //�������·����һ��Ҫ��������ļ�������������ļ�����չ����  
	            String savepath = new String ("E:\\���ǲ���text\\"+filename);    
	              
	            //����Word����  
	            ActiveXComponent app = new ActiveXComponent("Word.Application");          
	            //���������ļ�������ļ���·��  
	            String inFile = path;  
	            String tpFile = savepath;  
	            //����word���ɼ�  
	            app.setProperty("Visible", new Variant(false));  
	            //��䲻��  
	            Object docs = app.getProperty("Documents").toDispatch();  
	            //�������doc�ĵ�  
	            Object doc = Dispatch.invoke((Dispatch) docs,"Open", Dispatch.Method, new Object[]{inFile,new Variant(false), new Variant(true)}, new int[1]).toDispatch();  
	              
	            //����ļ�, ����Variant(n)����ָ�����Ϊ���ļ����ͣ������������������  
	            Dispatch.invoke((Dispatch) doc,"SaveAs", Dispatch.Method, new Object[]{tpFile,new Variant(2)}, new int[1]);  
	            //���Ҳ����  
	            Variant f = new Variant(false);  
	            //�رղ��˳�  
	            Dispatch.call((Dispatch) doc, "Close", f);  
	            app.invoke("Quit", new Variant[] {});  
	            System.out.println("ת����ϡ�");  
	        }
	    }
}
