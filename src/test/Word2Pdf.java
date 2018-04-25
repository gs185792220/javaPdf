package test;


import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

public class Word2Pdf {
    public static void main(String args[]) {
        ActiveXComponent app = null;
        String wordFile = "d:/word/test.docx";
       String pdfFile = "d:/word/testpdf.pdf";
       System.out.println("开始转换...");
       // 开始时间
       long start = System.currentTimeMillis();  
       try {
        // 打开word
        app = new ActiveXComponent("Word.Application");
        //app.setProperty("Visible", false);
        Dispatch documents = app.getProperty("Documents").toDispatch();
        System.out.println("打开文件: " + wordFile);
        // 打开文档
        Dispatch document = Dispatch.call(documents, "Open", wordFile, false, true).toDispatch();
        File target = new File(pdfFile);  
         if (target.exists()) {  
            target.delete();
         }
        System.out.println("另存为: " + pdfFile);
        Dispatch.call(document, "SaveAs", pdfFile, 17);
        Dispatch.call(document, "Close", false);
        long end = System.currentTimeMillis();
        System.out.println("转换成功，用时：" + (end - start) + "ms");
       }catch(Exception e) {
        System.out.println("转换失败"+e.getMessage());
       }finally {
        app.invoke("Quit", 0);
       }
    }
}
