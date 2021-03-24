package com.kiruah.poi2cc;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.Closeable;
import java.io.IOException;
import java.io.ObjectOutputStream;
import java.io.Serializable;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

public class Poi2ccUtil {

	public static boolean isEmptyString(CharSequence text) {

		if (text == null || text.length() == 0) {
			return true;
		}

		return false;
	}

	public static boolean equalString(CharSequence text1, CharSequence text2) {

		if (text1 == null && text2 == null) {
			return true;
		} else if (text1 != null && text2 == null) {
			return false;
		} else if (text1 == null && text2 != null) {
			return false;
		}

		if (text1.equals(text2) == true) {
			return true;
		}

		return false;
	}

	/**
	 * Map構造からEntry<K, V>のイテレータを取得します。
	 *
	 * <p>
	 * 使用例は下記の通りとなります。
	 * </p>
	 * <pre>
	 * Iterator<Entry<String, Integer>> iterator = Chore.getIterator(new HashMap<String, Integer>());
	 * </pre>
	 *
	 * @param <K> Mapのキー型
	 * @param <V> Mapの値型
	 * @param map Map情報
	 * @return Entryイテレータ
	 */
	public static <K, V> Iterator<Entry<K, V>> getIterator(Map<K, V> map) {

		Set<Entry<K, V>> set = map.entrySet();
		Iterator<Entry<K, V>> iterator = set.iterator();

		return iterator;
	}

	public static boolean isAsciiAlpha(char c) {

		if (('a' <= c && c <= 'z') || ('A' <= c && c <= 'Z')) {
			return true;
		}

		return false;
	}

	/**
	 * ファイル系APIのクローズ処理
	 *
	 * @param closeableArray クローズ対象のIOインスタンス
	 */
	public static void close(Closeable... closeableArray) {

		if (closeableArray == null || closeableArray.length == 0) {
			return;
		}

		for (Closeable c : closeableArray) {
			try {
				c.close();
			} catch (Exception e) {
				// ignore
			}
		}
	}

	@SuppressWarnings("unchecked")
	public static <T extends Serializable> T cloneObject(T object) {

		if (object == null) {
			return null;
		}

		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		ObjectOutputStream oos = null;

		try {
			oos = new ObjectOutputStream(baos);
			oos.writeObject(object);
		} catch (IOException e) {
			throw new Poi2ccRuntimeException("Object Serialization Error.", e);
		} finally {
			close(oos);
		}

		byte[] objectData = baos.toByteArray();
		ByteArrayInputStream bais = new ByteArrayInputStream(objectData);
		CustomObjectInputStream cois = null;

		try {
			cois = new CustomObjectInputStream(bais, object.getClass().getClassLoader());

			T cloneObject = (T) cois.readObject();
			return cloneObject;
		} catch (Exception e) {
			throw new Poi2ccRuntimeException("Object Clone Error.", e);
		} finally {
			close(cois);
		}
	}
}
