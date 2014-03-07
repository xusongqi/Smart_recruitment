
/*
 * 	transform .doc into .html
 *	xusongqi
 *	2013.7.26
 * */
package wordToHtml;

import com.jacob.com.*;  
import com.jacob.activeX.*;  
import java.io.*;


public class WordToHtml  
{    
    public static void main(String[] args)  
    {  
        //ָ����ת���ļ�������·���� ���������ͼ�ǰ�pdfתΪtxt  
        String path = new String("E:\\word2.doc");  
        //����·�������ļ�����  
        File docFile=new File(path);  
        //��ȡ�ļ�����������չ����  
        String filename=docFile.getName();  
        //���˵��ļ����е���չ��  
        int filenamelength = filename.length();  
        int dotposition=filename.indexOf(".");  
        filename=filename.substring(0,dotposition);  
          
        //�������·����һ��Ҫ��������ļ�������������ļ�����չ����  
        String savepath = new String ("E:\\"+filename);    
          
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

