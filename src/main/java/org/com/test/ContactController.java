
package org.com.test;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class ContactController {
 
    @RequestMapping("/dashboard")
     public String dashboard(Model model) {
         Contact contact = new Contact();
         model.addAttribute("contact", contact);
         return "dashboard";
    }

    @PostMapping("/saveContact")
    public String saveContact(@ModelAttribute("contact") Contact contact) {
        System.out.println("Contact Form:-->"+contact.getName());
        System.out.println("Contact Form:-->"+contact.getMessage());
        System.out.println("Contact Form:-->"+contact.getEmail());
        ExcelFileUtil.writeData(contact);
        return "dashboard";
    }
}
