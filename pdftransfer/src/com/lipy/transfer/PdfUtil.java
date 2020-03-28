package com.lipy.transfer;

import com.aspose.cells.Workbook;
import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.*;
import java.util.ArrayList;



/**
 * itext  转PDF 工具类
 * @author sunkuang
 *
 */
public class PdfUtil {
    public static File Pdf(ArrayList<String> imageUrllist,
                           String mOutputPdfFileName) {
        //Document doc = new Document(PageSize.A4, 20, 20, 20, 20); // new一个pdf文档
        com.itextpdf.text.Document doc = new com.itextpdf.text.Document();
        try {

            PdfWriter
                    .getInstance(doc, new FileOutputStream(mOutputPdfFileName)); // pdf写入
            doc.open();// 打开文档
            for (int i = 0; i < imageUrllist.size(); i++) { // 循环图片List，将图片加入到pdf中
                doc.newPage(); // 在pdf创建一页
                Image png1 = Image.getInstance(imageUrllist.get(i)); // 通过文件路径获取image
                float heigth = png1.getHeight();
                float width = png1.getWidth();
                int percent = getPercent2(heigth, width);
                png1.setAlignment(Image.MIDDLE);
                png1.scalePercent(percent + 3);// 表示是原来图像的比例;
                doc.add(png1);
            }
            doc.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        File mOutputPdfFile = new File(mOutputPdfFileName); // 输出流
        if (!mOutputPdfFile.exists()) {
            mOutputPdfFile.deleteOnExit();
            return null;
        }
        return mOutputPdfFile; // 反回文件输出流
    }

    public static int getPercent(float h, float w) {
        int p = 0;
        float p2 = 0.0f;
        if (h > w) {
            p2 = 297 / h * 100;
        } else {
            p2 = 210 / w * 100;
        }
        p = Math.round(p2);
        return p;
    }

    public static int getPercent2(float h, float w) {
        int p = 0;
        float p2 = 0.0f;
        p2 = 530 / w * 100;
        p = Math.round(p2);
        return p;
    }


    /**
     * 图片文件转PDF
     * @param filepath
     * @param request
     * @return
     */
    public static String imgOfPdf(String filepath) {
        boolean result = false;
        String pdfUrl = "E:\\test\\transferfgo.pdf";
        String fileUrl = "E:\\test\\transferfgo.pdf";
        try {
            result = getLicense();
            if (result == true) {
                ArrayList<String> imageUrllist = new ArrayList<>();
                imageUrllist.add("E:\\test\\fgo.jpg");
                File file = PdfUtil.Pdf(imageUrllist, pdfUrl);// 生成pdf
                file.createNewFile();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return fileUrl;
    }


	public static void doc2pdf(String Address, String outPath) {
		if (!getLicense()) { // 验证License 若不验证则转化出的pdf文档会有水印产生
			return;
		}
		try {
			long old = System.currentTimeMillis();
			File file = new File(outPath); // 新建一个空白pdf文档
			FileOutputStream os = new FileOutputStream(file);
			Document doc = new Document(Address); // Address是将要被转化的word文档
			doc.save(
					os,
					SaveFormat.PDF);// 全面支持DOC, DOCX, OOXML, RTF HTML,
			// OpenDocument, PDF, EPUB, XPS, SWF
			// 相互转换
			long now = System.currentTimeMillis();
			System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒"); // 转化用时
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

    public static String exceOfPdf(String filePath, String outPath) {
        if (!getLicense1()) {          // 验证License 若不验证则转化出的pdf文档会有水印产生
            return "PDF格式转化失败";
        }
        try {
            //long old = System.currentTimeMillis();

            //文件操作
            File file = new File(outPath); // 新建一个空白pdf文档
            FileOutputStream os = new FileOutputStream(file);
            Workbook wb = new Workbook(filePath);// 原始excel路径
            FileOutputStream fileOS = new FileOutputStream(file);
            wb.save(fileOS, com.aspose.cells.SaveFormat.PDF);
            fileOS.close();
            // long now = System.currentTimeMillis();
            //System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒");  //转化用时
            return outPath;

        } catch (Exception e) {
            e.printStackTrace();
        }
        return "PDF格式转化失败";
    }

    public static String pptOfpdf(String filePath, String outPath) {
        // 验证License
        if (!getLicense2()) {
            return "PDF格式转化失败";
        }
        try {
            long old = System.currentTimeMillis();
            //File file = new File("C:/Program Files (x86)/Apache Software Foundation/Tomcat 7.0/webapps/generic/web/file/pdf1.pdf");// 输出pdf路径
            //com.aspose.slides.Presentation pres = new  com.aspose.slides.Presentation(Address);//输入pdf路径

            //文件操作
            File file = new File(outPath); // 新建一个空白pdf文档
            com.aspose.slides.Presentation pres = new  com.aspose.slides.Presentation(filePath);//输入pdf路径

            FileOutputStream fileOS = new FileOutputStream(file);
            pres.save(fileOS, com.aspose.slides.SaveFormat.Pdf);
            fileOS.close();

            long now = System.currentTimeMillis();
            System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒\n\n" + "文件保存在:" + file.getPath()); //转化过程耗时
            return outPath;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "PDF格式转化失败";
    }
    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = PdfUtil.class.getClassLoader()
                    .getResourceAsStream("license.xml"); // license.xml应放在..\WebRoot\WEB-INF\classes路径下
//            is = new FileInputStream("E:\\project\\pdftransfer\\license.xml");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;

        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }


    public static boolean getLicense1() {
        boolean result = false;
        try {

            InputStream is = PdfUtil.class.getClassLoader()
                    .getResourceAsStream("license.xml"); // license.xml应放在..\WebRoot\WEB-INF\classes路径下
            com.aspose.cells.License aposeLic = new com.aspose.cells.License();
            aposeLic.setLicense(is);
            result = true;

        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public static boolean getLicense2() {
        boolean result = false;
        try {
            InputStream is = PdfUtil.class.getClassLoader()
                    .getResourceAsStream("license.xml"); // license.xml应放在..\WebRoot\WEB-INF\classes路径下
            com.aspose.slides.License aposeLic = new com.aspose.slides.License();
            aposeLic.setLicense(is);
            result = true;

        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public static void main(String[] args) {
//        imgOfPdf("");
//        doc2pdf("e:\\test\\1.docx", "E:\\test\\1d.pdf");
//        exceOfPdf("E:\\test\\信.xls", "E:\\test\\1x.pdf");
        pptOfpdf("E:\\test\\1 MongoDB综述（一）.pptx", "E:\\test\\ptp.pdf");
    }

  /*
   * 因为TXT 可以直接用上面的  DOC 方法 转    暂时 不用这个
   * public static void textOfpdf(String filePath,HttpServletRequest request) throws DocumentException, IOException {

    	String text =  request.getSession().getServletContext().getRealPath("\\" + filePath);
      	String pdf = filePath.substring(0, filePath.lastIndexOf("."));

    	BaseFont bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
    	Font FontChinese = new Font(bfChinese, 12, Font.NORMAL);
    	FileOutputStream out = new FileOutputStream(pdf);
    	Rectangle rect = new Rectangle(PageSize.A4.rotate());
    	com.itextpdf.text.Document doc = new com.itextpdf.text.Document(rect);
    	PdfWriter writer = PdfWriter.getInstance(doc, out);
    	doc.open();
    	Paragraph p = new Paragraph();
    	p.setFont(FontChinese);

    	BufferedReader read = new BufferedReader(new FileReader(text));

    	String line = read.readLine();
    	while(line != null){
    	System.out.println(line);
    	p.add(line+"\n");
    	line = read.readLine();
    	}
    	read.close();
    	doc.add(p);
    	doc.close();

    }*/
}