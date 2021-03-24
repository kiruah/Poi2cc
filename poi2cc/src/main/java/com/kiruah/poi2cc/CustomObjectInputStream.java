package com.kiruah.poi2cc;

import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.ObjectStreamClass;
import java.util.HashMap;
import java.util.Map;

public class CustomObjectInputStream extends ObjectInputStream {

	private static final Map<String, Class<?>> primitiveTypeMap = new HashMap<>();

	private ClassLoader cl = null;

	static {

		Class<?>[] classArray = new Class<?>[] {byte.class, char.class, short.class, int.class, long.class, float.class, double.class, boolean.class, void.class};

		for (Class<?> c : classArray) {
			primitiveTypeMap.put(c.getName(), c);
		}
	}

	public CustomObjectInputStream(InputStream is, ClassLoader cl) throws IOException {

		super(is);
		this.cl = cl;
	}

	@Override
	protected Class<?> resolveClass(ObjectStreamClass desc) throws IOException, ClassNotFoundException {

		String name = desc.getName();

		try {
			return Class.forName(name, false, cl);
		} catch (ClassNotFoundException e) {
		}

		try {
			return Class.forName(name, false, Thread.currentThread().getContextClassLoader());
		} catch (ClassNotFoundException e) {
			Class<?> c = primitiveTypeMap.get(name);
			if (c != null) {
				return c;
			}

			throw e;
		}
	}
}
