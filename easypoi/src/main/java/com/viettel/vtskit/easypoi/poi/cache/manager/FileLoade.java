package com.viettel.vtskit.easypoi.poi.cache.manager;

import com.viettel.vtskit.easypoi.poi.common.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

class FileLoade {

	private static final Logger LOGGER = LoggerFactory.getLogger(FileLoade.class);

	public byte[] getFile(String url) {
		FileInputStream fileis = null;
		ByteArrayOutputStream baos = null;
		try {
			// First use the absolute path to query, and then query the relative path
			try {
				fileis = new FileInputStream(url);
			} catch (FileNotFoundException e) {
				String path = PoiPublicUtil.getWebRootPath(url);
				fileis = new FileInputStream(path);
			}
			baos = new ByteArrayOutputStream();
			byte[] buffer = new byte[1024];
			int len;
			while ((len = fileis.read(buffer)) > -1) {
				baos.write(buffer, 0, len);
			}
			baos.flush();
			return baos.toByteArray();
		} catch (FileNotFoundException e) {
			LOGGER.error(e.getMessage(), e);
		} catch (IOException e) {
			LOGGER.error(e.getMessage(), e);
		} finally {
			try {
				if (fileis != null)
					fileis.close();
				if (fileis != null)
					baos.close();
			} catch (IOException e) {
				LOGGER.error(e.getMessage(), e);
			}
		}
		LOGGER.error(fileis + "The path file is not found, please query");
		return null;
	}

}
