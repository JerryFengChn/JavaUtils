package com.yumaolin.util.PDFBoxUtil;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfWriter;

public class ITextUtil {
	
	/**
	 * <p>创建PDF文件</p>
	 * @param filePath
	 */
	public void createPDFFile(String filePath) {
		try{
			Document doc = new Document();
			String fontPath = "C:\\Windows\\Fonts\\MSYHMONO.ttf";
			PdfWriter.getInstance(doc, new FileOutputStream(filePath));
			doc.open();
			BaseFont bfChinese = BaseFont.createFont(fontPath,BaseFont.IDENTITY_H,BaseFont.EMBEDDED);
			Font fontChinese = new Font(bfChinese, 12, Font.NORMAL, BaseColor.GREEN);
			Paragraph pf = new Paragraph("protocolNo|custIdenNo|custName", fontChinese);
			//pf.add(new Paragraph("我们的家好大的家", fontChinese));
			//pf.add(new Paragraph("sfsfsf"));
			doc.add(pf);
			doc.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void parsePDFFile(String oldPath,String newPath,Map<String,String> value) throws IOException {
		Path out = Paths.get(newPath);
		Files.deleteIfExists(out);
		out = Paths.get(newPath);
		PdfReader reader = null;
		PdfStamper ps = null;
		try {
			reader = new PdfReader(oldPath); // 模版文件目录
			ps = new PdfStamper(reader, new FileOutputStream(out.toFile())); // 生成的输出流
			BaseFont bf = BaseFont.createFont("STSong-Light","UniGB-UCS2-H",BaseFont.NOT_EMBEDDED);
			AcroFields s =  ps.getAcroFields();
			if(value != null){
				for(String mapKey : value.keySet()){
					s.setFieldProperty(mapKey,"textfont",bf,null);
					s.setField(mapKey,value.get(mapKey));
				}
			}
			/*//设置文本域表单的字体
			s.setFieldProperty("custName","textfont",bf,null);
			s.setFieldProperty("merchantName","textfont",bf,null);
			s.setFieldProperty("1", "textfont",bf,null);
			s.setFieldProperty("2", "textfont",bf,null);
			s.setFieldProperty("bankType","textfont",bf,null);
			s.setFieldProperty("custBankCardNo","textfont",bf,null);
			s.setFieldProperty("custPhoneNo","textfont",bf,null);
			s.setFieldProperty("Signature1","textfont",bf,null);
			s.setFieldProperty("Signature2","textfont",bf,null);
			s.setFieldProperty("signDate","textfont",bf,null);
			//编辑文本域表单的内容
			s.setField("custName","测试客户");//授权人
			s.setField("merchantName", "测试商户");//被授权人
			s.setField("1","2018-07-31");//签署业务合同时间
			s.setField("2","代购协议");//合同名称
			s.setField("bankType","招商银行");//授权人缴费开户银行
			s.setField("custBankCardNo","62284801245784");//银行卡账号
			s.setField("custPhoneNo","18684320213");//授权人联系手机
			s.setField("Signature1","测试客户1");//授权人签名
			s.setField("Signature2","测试商户2");//被授权人签名
			s.setField("signDate","2018-07-31");//签订日期*/			
			ps.setFormFlattening(true);//如果为false那么生成的PDF文件还能编辑，一定要设为true	
		} catch (Exception e) {
			//logger.warn("生成合同预览文件出错", e);
			e.printStackTrace();
		}finally{
			if(ps != null){
				try {
					ps.close();
				} catch (DocumentException e) {
				}
			}
			if(reader != null){
				reader.close();
			}
		}
		//logger.info("商户下载入网合同,全路径[{}]", out);
	}
	
	public static void main(String[] args) throws IOException, DocumentException {
		ITextUtil textUtil = new ITextUtil();
		//textUtil.createPDFFile("D:\\test.PDF");
		//PDFReplacer textReplacer = new PDFReplacer("D://test.PDF");  
	    //textReplacer.replaceText("protocolNo", "小白");  
	    //textReplacer.replaceText("本科", "社会大学");  
	    //textReplacer.replaceText("0755-29493863", "15112345678");  
	    //textReplacer.toPdf("D://newTest.PDF");  
		Map<String,String> value = new HashMap<String,String>();
		value.put("custName","测试客户");
		value.put("merchantName", "测试商户");
		value.put("1","2018/07/31");
		value.put("2","代购协议");
		value.put("custBankCardNo","62284801245784");
		value.put("custPhoneNo","18684320213");
		value.put("Signature1","测试客户1");
		value.put("Signature2","测试商户2");
		value.put("signDate","2018/07/31");
		value.put("bankType","招商银行");
		textUtil.parsePDFFile("D:\\委托代扣款协议测试（旧）.pdf","D:\\委托代扣款协议测试（新）.pdf",value);
	}
}
