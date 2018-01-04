package com.yumaolin.util.MessageAttestation;

import java.nio.charset.Charset;
import java.security.PrivateKey;
import java.security.cert.X509Certificate;
import org.apache.commons.codec.binary.Base64;
import org.springframework.util.Assert;


public class KeystoreSignProvider implements SignProvider {
	private final X509Certificate x509Certificate;
	private final PrivateKey privateKey;

	public KeystoreSignProvider(String kyeStoreType, String keyStorePath, char[] keyStorePassword, String alias,
			char[] keyPassword) throws Exception {
		Assert.hasText(keyStorePath, "客户端keystore路径不能为空");
		Assert.notNull(keyStorePassword, "keystore密码不能为空");
		this.x509Certificate = ((X509Certificate) EncryptionUtils.getCertificate(kyeStoreType, keyStorePath,
				keyStorePassword, alias));

		this.privateKey = EncryptionUtils.getPrivateKeyByKeyStore(kyeStoreType, keyStorePath, keyStorePassword, alias,
				keyPassword);
	}

	public String sign(byte[] plaintext, Charset charset) throws Exception {
		byte[] signatureInfo = EncryptionUtils.signByX509Certificate(plaintext, this.x509Certificate, this.privateKey);
		return Base64.encodeBase64String(signatureInfo);
	}
}
