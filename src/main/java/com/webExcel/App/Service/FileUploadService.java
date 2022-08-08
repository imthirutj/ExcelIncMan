package com.webExcel.App.Service;

import java.io.*;
import java.nio.file.Path;

import javax.servlet.http.HttpServletRequest;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class FileUploadService {

	


    
	public void uploadFile(MultipartFile file) throws IllegalStateException, IOException {
	
		
	
		file.transferTo(new File("ftp://155.254.244.30/www.tjapplications.somee.com/JAVA_EXCEL/"+"file1"));
	}
}
