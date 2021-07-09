package com.cssca.cloud;

import org.apache.http.entity.ContentType;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.UUID;

@Controller
public class ConvertController {
    @Resource
    DocConvertService docConvertService;

    @RequestMapping(value = "word2Pdf_openoffice",method = RequestMethod.POST)
    public void word2PdfByOpenOffice(HttpServletRequest request, HttpServletResponse response) throws Exception{
        InputStream in=request.getInputStream();
        String fileName= UUID.randomUUID().toString();
        response.setHeader("Content-Type", ContentType.APPLICATION_OCTET_STREAM.toString());
        response.setHeader("Content-Disposition", "attachment; filename="+fileName);
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        OutputStream out=response.getOutputStream();
        docConvertService.office2PdfByOpenOffice(in,out);
        out.flush();
        out.close();
    }

    @RequestMapping(value = "word2Pdf_documents4j",method = RequestMethod.POST)
    public void word2PdfByMSOffice(HttpServletRequest request, HttpServletResponse response) throws Exception{
        String fileName= UUID.randomUUID().toString()+".doc";
        response.setHeader("Content-Type", ContentType.APPLICATION_OCTET_STREAM.toString());
        response.setHeader("Content-Disposition", "attachment; filename="+fileName);
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        docConvertService.office2PdfByMSOffice(request.getInputStream(),response.getOutputStream());
    }

    @RequestMapping(value = "word2Pdf_libreoffice",method = RequestMethod.POST)
    public void word2PdfByLibreOffice(HttpServletRequest request, HttpServletResponse response) throws Exception{
        String fileName= UUID.randomUUID().toString();
        response.setHeader("Content-Type", ContentType.APPLICATION_OCTET_STREAM.toString());
        response.setHeader("Content-Disposition", "attachment; filename="+fileName);
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        docConvertService.office2PdfByByLibOffice(request.getInputStream(),response.getOutputStream(),fileName);
    }

//    @PostConstruct
//    public void test(){
//        String inputWord="D:\\workspace\\TACS\\tacs\\let\\letWeb\\src\\main\\webapp\\static\\let\\letterTemplate\\REGL10_ENGLISH.docx";
//        String outputFile="D:\\test2.pdf";
////        docConvertService.office2PDF(inputWord,outputFile);
//        File file=new File(inputWord);
//        try (FileInputStream docxInputStream = new FileInputStream(inputWord);
//             OutputStream outputStream = new FileOutputStream(outputFile)){
//            IConverter  converter = LocalConverter.builder().baseFolder(file.getAbsoluteFile())
//                    .processTimeout(5L, TimeUnit.MINUTES)
//                    .workerPool(15,30,10,TimeUnit.MINUTES).build();
//            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
//            outputStream.close();
//            System.out.println("success");
//        }catch (Exception e){
//            e.printStackTrace();
//        }
//    }
}
