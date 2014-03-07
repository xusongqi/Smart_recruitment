package traversalFiles;

import java.io.File;
import java.io.FileFilter;

public class TraversalFiles {
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
	    }
}
