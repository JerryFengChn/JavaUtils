package com.yumaolin.util.MessageAttestation;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.security.Signature;
import java.security.cert.Certificate;
import java.security.cert.CertificateFactory;
import java.security.cert.X509Certificate;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.codec.binary.Base64;
/**
 * 报文验签工具类
 */
public class SignUtil {

	
	private final static String charset = "UTF-8";

	private static Map<String, String> param = null;
	private static String signatureInfo = "";

	public static void main(String[] args) {
		String caPath = "E:\\downFile\\10-cert\\";
		// 请求报文加签
		sign(caPath + "pfx.pfx", "123456");
		// 同步报文验签
		verifySign(caPath + "pfx.cer");
	}

	/**
	 * 同步报文验证签约
	 * 
	 * @param cerPath
	 */
	public static void verifySign(String cerPath) {
		try {
			Map<String, String> parameters = paramsFilter(param);
			String createLinkString = createLinkString(parameters);
			byte[] datas = createLinkString.getBytes(charset);
			// 得到的报文中的签名
			String signMsg = signatureInfo;
			byte[] sign = Base64.decodeBase64(signMsg);
			boolean verifySign = verifySign(datas, sign, cerPath);
			System.out.println("验签名结果：" + verifySign);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void verifySign(String cerPath, Map<String, String> param, String signatureInfo) {
		try {
			Map<String, String> parameters = paramsFilter(param);
			String createLinkString = createLinkString(parameters);
			byte[] datas = createLinkString.getBytes(charset);
			// 得到的报文中的签名
			String signMsg = signatureInfo;
			byte[] sign = Base64.decodeBase64(signMsg);
			boolean verifySign = verifySign(datas, sign, cerPath);
			System.out.println("验签名结果：" + verifySign);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 请求报文签约
	 * 
	 * @param pfxPath
	 * @param psw
	 */
	private static void sign(String pfxPath, String psw) {
		try {
			SignProvider signProvider = new KeystoreSignProvider("PKCS12", pfxPath, psw.toCharArray(), null,
					psw.toCharArray());
			param = new HashMap<String, String>();
			param.put("orderNo", "201801031509002");
			param.put("status", "1");
			param.put("callerIp", "10.37.20.46");
			param.put("signatureAlgorithm", "RSA");
			param.put("reconStatus", "1");
			param.put("settlementAmount", "1");
			param.put("checkDate", "20180103");
			param.put("language", "zh_CN");
			param.put("channelNo", "4100002085065666279205266395136");
			Map<String, String> signParameters = signParameters(signProvider, param);
			signatureInfo = signParameters.get("signatureInfo");
			System.out.println(signatureInfo);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static Map<String, String> signParameters(SignProvider signProvider,final Map<String, String> originalParameters) throws Exception {
		final Map<String, String> parameters = paramsFilter(originalParameters);
		final String prestr = createLinkString(parameters);
		System.out.println("prestr:" + prestr);
		Charset encoding = Charset.forName(charset);
		String encodeBase64String = signProvider.sign(prestr.getBytes(charset), encoding);
		parameters.put("signatureInfo", encodeBase64String);
		parameters.put("signatureAlgorithm", "RSA");
		return parameters;
	}

	/**
	 * 除去参数中的空值和签名参数
	 * 
	 * @param parameters 签名参数组
	 * @return 去掉空值与签名参数后的新签名参数组
	 */
	private static Map<String, String> paramsFilter(final Map<String, String> parameters) {
		if (parameters == null || parameters.size() == 0) {
			return new HashMap<String, String>();
		}
		final Map<String, String> result = new HashMap<String, String>(parameters.size());
		String value = null;
		for (final String key : parameters.keySet()) {
			value = parameters.get(key);
			if (value == null || "".equals(value) || key.equalsIgnoreCase("signatureAlgorithm")
					|| key.equalsIgnoreCase("signatureInfo")) {
				continue;
			}
			result.put(key, value);
		}
		return result;
	}

	/**
	 * 对请求参数排序，并按照接口规范中所述"参数名=参数值"的模式用"&"字符拼接成字符串
	 * 
	 * @param params 需要排序并参与字符拼接的参数
	 * @return 拼接后字符串
	 */
	private static String createLinkString(final Map<String, String> params) {

		final List<String> keys = new ArrayList<String>(params.keySet());
		Collections.sort(keys);
		final StringBuilder sb = new StringBuilder();
		String key = null;
		String value = null;
		for (int i = 0; i < keys.size(); i++) {
			key = keys.get(i);
			value = params.get(key);
			sb.append(key).append("=").append(value);
			// 最后一组参数,结尾不包括'&'
			if (i < keys.size() - 1) {
				sb.append("&");
			}
		}
		return sb.toString();
	}

	/**
	 * <p> 验证签名,从证书中获取公钥来验证签名是否正确 </p>
	 * 
	 * @param data 传输的数据
	 * @param sign 对传输数据的签名
	 * @param certificateContent 证书内容的2进制形式
	 * @return
	 * @throws Exception
	 * @author 汪一鸣
	 */
	public static boolean verifySign(byte[] data, byte[] sign, String certificateContent) throws Exception {
		X509Certificate certificate = (X509Certificate) getCertificate(certificateContent);
		Signature signature = Signature.getInstance(certificate.getSigAlgName());
		// 由证书初始化签名,使用了证书中的公钥
		signature.initVerify(certificate);
		signature.update(data);
		return signature.verify(sign);
	}

	/**
	 * <p> 从证书文件读取证书.'.crt'和'.cer'文件都可以读取 .cer是IE导出的公钥证书（der格式） </p>
	 * 
	 * @param certificatePath 证书文件路径:可以直接加载指定的文件,例如"file:C:/kft.cer"
	 * @throws Exception
	 * @author 汪一鸣
	 */
	private static Certificate getCertificate(String certificatePath) throws Exception {
		File certificate = new File(certificatePath);
		if (certificate == null || (certificate.exists() && certificate.isDirectory())) {
			throw new IllegalArgumentException("certificatePath[" + certificatePath + "]必须是一个已经存在的文件,不能为空,且不能是一个文件夹");
		}
		InputStream inputStream = null;
		try {
			inputStream = new FileInputStream(certificate);
			// 实例化证书工厂
			CertificateFactory cf = CertificateFactory.getInstance("X.509");
			Certificate cert = cf.generateCertificate(inputStream);
			return cert;
		} finally {
			if (inputStream != null) {
				inputStream.close();
			}
		}
	}
}
