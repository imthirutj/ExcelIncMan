package com.webExcel.App.Controller;

import java.io.File;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class DeleteFiles {

	
	@GetMapping("DeleteFiles")
	public String Delete(Model model) {

        File file1
            = new File(".\\src\\main\\resources\\webapp\\uploads\\IncMan.xlsx");
 
        File file2
        = new File(".\\src\\main\\resources\\webapp\\uploads\\IncidentMan.xlsx");

        
        Boolean bfile1=file1.delete();
        Boolean bfile2=file2.delete();
        if (bfile1) {
            System.out.println("IncMan.xlsx deleted successfully");
            model.addAttribute("f1msg", "IncMan.xlsx deleted successfully");
        }
        if(bfile2) {
            System.out.println("IncidentMan.xlsx deleted successfully");
            model.addAttribute("f2msg", "IncidentMan.xlsx deleted successfully");
        }
       if(!bfile1 && !bfile2) {
            System.out.println("Failed to delete the file");
            model.addAttribute("failmsg", "Failed to deleted successfully");
        }
        
        
		return "showFiles";
	}
}
