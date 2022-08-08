package com.webExcel.App.Controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.net.SocketException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import javax.servlet.http.HttpServletRequest;

import org.apache.commons.net.ftp.FTPClient;
import org.apache.commons.net.ftp.FTPConnectionClosedException;
import org.apache.tomcat.util.ExceptionUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;


import com.webExcel.App.Service.FileUploadService;

@Controller
public class FileUploadController {
  public static String uploadDirectory = System.getProperty("user.dir")+"/src/main/resources/webapp/uploads";
  @GetMapping("/upload")
  public String UploadPage(Model model) {
	  return "uploadview";
  }
  @PostMapping("/uploadIncMan")
  public String upload(Model model,@RequestParam("file") MultipartFile[] files) throws IOException, IOException {
		
		
		  StringBuilder fileNames = new StringBuilder();
		  for (MultipartFile file:files) { Path fileNameAndPath = Paths.get(uploadDirectory,
		  file.getOriginalFilename());
		  fileNames.append(file.getOriginalFilename()+" "); try {
		  Files.write(fileNameAndPath, file.getBytes()); model.addAttribute("msg",
		  "Successfully uploaded files "+fileNames.toString()); } catch (IOException e)
		  { e.printStackTrace();
		  
		  StackTraceElement[] stack = new Exception().getStackTrace(); String theTrace
		  = ""; for(StackTraceElement line : stack) { theTrace += line.toString()+"\n";
		  model.addAttribute("msg",theTrace); }
		  
		  } }
		 
	  
	  
	  
	  
		/*
		 * FTPClient client = new FTPClient(); //FileInputStream fis = null; boolean
		 * result; try { client.connect("files.000webhost.com"); result =
		 * client.login("tjproductionz", "121212@Tj");
		 * 
		 * if (result == true) { System.out.println("Successfully logged in!");
		 * 
		 * 
		 * 
		 * // File file = new File("D:/Files/sampleftp.txt"); String testName =
		 * "/Files/"+files.getOriginalFilename();
		 * 
		 * File file = new File( files.getOriginalFilename()); files.transferTo(file);
		 * 
		 * FileInputStream fis = new FileInputStream(file); // Upload file to the ftp
		 * server result = client.storeFile(testName, fis);
		 * 
		 * 
		 * if (result == true) { System.out.println("File is uploaded successfully"); }
		 * else { System.out.println("File uploading failed"); } client.logout(); } else
		 * { System.out.println("Login Fail!");
		 * 
		 * }
		 * 
		 * } catch (FTPConnectionClosedException e) { e.printStackTrace(); } finally {
		 * try { client.disconnect(); } catch (FTPConnectionClosedException e) {
		 * System.out.println(e); } }
		 */
	 

	  return "uploadstatusview";
  }
  
  
  
  
}
