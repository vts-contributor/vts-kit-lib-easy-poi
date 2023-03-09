package com.viettel.vtskit.easypoi.poi.cache.manager;

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;
public final class POICacheManager {

	private static final Logger LOGGER = LoggerFactory.getLogger(POICacheManager.class);

	private static LoadingCache<String, byte[]> loadingCache;

	static {
		loadingCache = CacheBuilder.newBuilder().expireAfterWrite(7, TimeUnit.DAYS).maximumSize(50).build(new CacheLoader<String, byte[]>() {
			@Override
			public byte[] load(String url) throws Exception {
				return new FileLoade().getFile(url);
			}
		});
	}

	public static InputStream getFile(String id) {
		try {
			// Complex data, prevent manipulation of original data
			byte[] result = Arrays.copyOf(loadingCache.get(id), loadingCache.get(id).length);
			return new ByteArrayInputStream(result);
		} catch (ExecutionException e) {
			LOGGER.error(e.getMessage(), e);
		}
		return null;
	}

}
