package com.yumaolin.util.MessageAttestation;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.security.KeyStore;
import java.security.PrivateKey;
import java.security.Signature;
import java.security.cert.Certificate;
import java.security.cert.X509Certificate;
import java.util.Collections;
import java.util.List;


public abstract class EncryptionUtils {
	public static final String CERTIFICATE_TYPE = "X.509";

	public static byte[] signByX509Certificate(byte[] data, X509Certificate x509Certificate, PrivateKey privateKey)
			throws Exception {
		Signature signature = Signature.getInstance(x509Certificate.getSigAlgName());

		signature.initSign(privateKey);
		signature.update(data);
		return signature.sign();
	}

	public static PrivateKey getPrivateKeyByKeyStore(String keystoreType, String keyStorePath, char[] keyStorePassword,
			String alias, char[] keyPassword) throws Exception {
		KeyStore ks = loadKeyStore(keyStorePath, keyStorePassword, keystoreType);
		if (alias == null) {
			List<String> aliases = Collections.list(ks.aliases());
			if (aliases.size() == 1) {
				alias = (String) aliases.get(0);
			} else {
				throw new IllegalArgumentException(
						"[Assertion failed] - this String argument[alias] must have text; it must not be null, empty, or blank");
			}
		}
		PrivateKey key = (PrivateKey) ks.getKey(alias, keyPassword);
		return key;
	}

	public static KeyStore loadKeyStore(String keyStorePath, char[] password, String keystoreType) throws Exception {
		KeyStore ks = KeyStore.getInstance(keystoreType == null ? KeyStore.getDefaultType() : keystoreType);
		File keystore = new File(keyStorePath);
		if ((keystore == null) || ((keystore.exists()) && (keystore.isDirectory()))) {
			throw new IllegalArgumentException("keystore[" + keyStorePath + "]");
		}
		InputStream is = null;
		try {
			is = new FileInputStream(keystore);
			ks.load(is, password);
			return ks;
		} finally {
			if (is != null) {
				is.close();
			}
		}
	}

	public static Certificate getCertificate(String keystoreType, String keyStorePath, char[] keyStorePassword,
			String alias) throws Exception {
		KeyStore ks = loadKeyStore(keyStorePath, keyStorePassword, keystoreType);
		if (alias == null) {
			List<String> aliases = Collections.list(ks.aliases());
			if (aliases.size() == 1) {
				alias = (String) aliases.get(0);
			} else {
				throw new IllegalArgumentException(
						"[Assertion failed] - this String argument[alias] must have text; it must not be null, empty, or blank");
			}
		}
		return ks.getCertificate(alias);
	}
}
