package com.sidoba.santamar.excelTest.POCExcelFormuleValeurLire.api;

import com.sidoba.santamar.excelTest.POCExcelFormuleValeurLire.util.LireExcelUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

@Controller
@RequestMapping("/api")
public class LireExcelControlleur {
    @GetMapping("/lire-excel")
    public ModelAndView PagePrincipale() {
        ModelAndView mav = new ModelAndView("html/lire-excel");
        return mav;
    }


    @PostMapping("/lire-excel/chargerFichier")
    public String traitementDuFichierExcel(@RequestParam("file") MultipartFile multipartFile, RedirectAttributes redirectAttributes) {
        String tableHTML = LireExcelUtil.creerTableValeurs(multipartFile);
        redirectAttributes.addFlashAttribute("tableHTML", tableHTML);

        return "redirect:/api/lire-excel-table";
    }

    @GetMapping("/lire-excel-table")
    public ModelAndView afficherInformationsExcel(@ModelAttribute("tableHTML") String tableHTML) {
        ModelAndView mav = new ModelAndView("html/lire-excel-table");

        if (tableHTML != null) {
            mav.addObject("tableHTML", tableHTML);
        }


        return mav;
    }
}
