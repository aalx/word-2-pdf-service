package com.cssca.cloud;

import com.artofsolving.jodconverter.DefaultDocumentFormatRegistry;
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.io.*;
import java.net.ConnectException;
import java.util.concurrent.TimeUnit;

@Service("docConvertService")
@ConfigurationProperties(prefix = "file.convert.openoffice")
@Slf4j
public class DocConvertService {
    /**
     * 记得修改自己的openoffice路径
     * 下面host_Str port_Str的一般都是默认 /opt/openoffice4 D:/OpenOffice 4
     */
//    private  String OpenOffice_HOME = "C:\\Program Files (x86)\\OpenOffice 4";
////    private  String host_Str = "127.0.0.1";
////    private  String port_Str = "8100";

    private String programPath;
    private String serverUrl;
    private String port;

    public String getProgramPath() {
        return programPath;
    }

    public void setProgramPath(String programPath) {
        this.programPath = programPath;
    }

    public String getServerUrl() {
        return serverUrl;
    }

    public void setServerUrl(String serverUrl) {
        this.serverUrl = serverUrl;
    }

    public String getPort() {
        return port;
    }

    public void setPort(String port) {
        this.port = port;
    }

    @PostConstruct
    public void cleanTemp(){
        try {
            String tempPath = PathUtils.getStaticPath() + PathUtils.getSeparator() + "temp";
            File baseFolder = new File(tempPath);
            if (baseFolder.exists()) {
                this.remove(baseFolder);
                log.debug("Delete temp file 【{}】.",tempPath);
            }
        }catch (Exception e){
            log.warn("Temp file delete fail.");
        }
    }

    public  void remove(File dir) {
        File files[] = dir.listFiles();
        for (int i = 0; i < files.length; i++) {
            if (files[i].isDirectory()) {
                remove(files[i]);
            } else {
                //删除文件
                log.debug("deleted  ::  " + files[i].toString());
                files[i].delete();
            }
        }
        //删除目录
        dir.delete();
        log.debug("deleted  ::  " + dir.toString());
    }
    /**
     * word文档转pdf文件
     *
     * @param sourceFile office文档绝对路径
     * @param destFile   pdf文件绝对路径
     * @return
     */
    public  int office2PdfByOpenOffice(String sourceFile, String destFile) {
        File inputFile = new File(sourceFile);
        if (!inputFile.exists()) {
            return -1; // 找不到源文件
        }

        // 如果目标路径不存在, 则新建该路径
        File outputFile = new File(destFile);
        if (!outputFile.getParentFile().exists()) {
            outputFile.getParentFile().mkdirs();
        }
        try {
            OpenOfficeConnection connection = initOpenOffice();
            if (!outputFile.exists()) {
                return convertFile(inputFile, outputFile, connection);
            }
            // 关闭连接和服务
            connection.disconnect();
            return -1;
        } catch (ConnectException e) {
            OpenOfficeConnection connection = initOpenOffice();
            try {
                return convertFile(inputFile, outputFile, connection);
            } catch (ConnectException e1) {
                e1.printStackTrace();
            }
        }
        return 1;
    }



    //                          下面是工具方法
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    /**
     * 初始化启动openoffice服务
     *
     * @throws IOException
     */
    private  OpenOfficeConnection initOpenOffice() {
        try {
            // 启动OpenOffice的服务
            String separator = PathUtils.getSeparator();
            String command = null;
            if (separator.equals("/")) {
                //linux系统
                command = this.programPath
                        + "/program/soffice -headless -accept=\"socket,host="
                        + this.serverUrl + ",port=" + this.port + ";urp;\" -nofirststartwizard";
            } else {
                command = this.programPath
                        + "/program/soffice.exe -headless -accept=\"socket,host="
                        + this.serverUrl + ",port=" + this.port + ";urp;\" -nofirststartwizard";
            }
            Runtime.getRuntime().exec(command);
            OpenOfficeConnection connection = new SocketOpenOfficeConnection(Integer.parseInt(this.port));
            return connection;
        } catch (Exception e) {
            log.warn("###\n启动openoffice服务错误",e);
        }
        return null;
    }



    private static int convertFile(File inputFile, File outputFile, OpenOfficeConnection connection) throws ConnectException {
        connection.connect();
        DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
        converter.convert(inputFile, outputFile);
        connection.disconnect();
        log.debug("****pdf转换成功，PDF输出：" + outputFile.getPath() + "****");
        return 0;

    }

    public  int office2PdfByOpenOffice(InputStream inputStream, OutputStream outputStream) {

        try {
            OpenOfficeConnection connection = initOpenOffice();
            int result=convertFile(inputStream, outputStream, connection);
            return result;
        } catch (ConnectException e) {
            OpenOfficeConnection connection = initOpenOffice();
            try {
                return convertFile(inputStream, outputStream, connection);
            } catch (ConnectException e1) {
                e1.printStackTrace();
            }
        }
        return 1;
    }


        private  int convertFile(InputStream inputString, OutputStream outputStream, OpenOfficeConnection connection) throws ConnectException {
            connection.connect();
            DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
            DefaultDocumentFormatRegistry formatReg = new DefaultDocumentFormatRegistry();
            DocumentFormat pdfFormat = formatReg.getFormatByFileExtension("pdf");
            DocumentFormat docFormat = formatReg.getFormatByFileExtension("doc");
            converter.convert(inputString,docFormat, outputStream,pdfFormat);
            connection.disconnect();
            return 0;

        }



    /**
     * 利用libreOffice将office文档转换成pdf
     * @param inputStream  目标文件地址
     * @param outputStream    输出文件夹
     * @return
     */
    public  boolean office2PdfByByLibOffice(InputStream inputStream,OutputStream outputStream,String fileName) throws Exception{
        File docFile=this.saveTempFile(inputStream,fileName+".doc");
        String tempPath=PathUtils.getStaticPath()+PathUtils.getSeparator()+"temp"+PathUtils.getSeparator();
        long start = System.currentTimeMillis();
        String command;
        boolean flag;
        String osName = System.getProperty("os.name");
        if (osName.contains("Windows")) {
            command = "cmd /c soffice --headless --invisible --convert-to pdf:writer_pdf_Export " + tempPath+fileName+".doc" + " --outdir " + tempPath;
        }else {
            command = "libreoffice --headless --invisible --convert-to pdf:writer_pdf_Export " + tempPath+fileName+".doc"  + " --outdir " + tempPath;
        }
        flag = executeLibreOfficeCommand(command);
        long end = System.currentTimeMillis();
        log.debug("用时:{} ms", end - start);
        File tempPdfFile=new File(tempPath+fileName+".pdf");
        if(tempPdfFile.exists()){
            try(FileInputStream _inputStream=new FileInputStream(tempPdfFile)){
                byte[] buf = new byte[1024];
                int len;
                while((len=_inputStream.read(buf))>0){
                    outputStream.write(buf,0,len);
                }
            }catch (Exception e){
                e.printStackTrace();
            }finally {
                outputStream.flush();
                outputStream.close();
            }
        }
        if(docFile.exists()){
            docFile.delete();
        }
        return flag;
    }


    /**
     * 执行command指令
     * @param command
     * @return
     */
    private  boolean executeLibreOfficeCommand(String command) {
        log.info("开始进行转换.......");
        Process process;// Process可以控制该子进程的执行或获取该子进程的信息
        try {
            log.debug("convertOffice2PDF cmd : {}", command);
            process = Runtime.getRuntime().exec(command);// exec()方法指示Java虚拟机创建一个子进程执行指定的可执行程序，并返回与该子进程对应的Process对象实例。
            // 下面两个可以获取输入输出流
//            InputStream errorStream = process.getErrorStream();
//            InputStream inputStream = process.getInputStream();
        } catch (IOException e) {
            log.error(" convertOffice2PDF {} error", command, e);
            return false;
        }
        int exitStatus = 0;
        try {
            exitStatus = process.waitFor();// 等待子进程完成再往下执行，返回值是子线程执行完毕的返回值,返回0表示正常结束
            // 第二种接受返回值的方法
            int i = process.exitValue(); // 接收执行完毕的返回值
            log.debug("i----" + i);
        } catch (InterruptedException e) {
            log.error("InterruptedException  convertOffice2PDF {}", command, e);
            return false;
        }
        if (exitStatus != 0) {
            log.error("convertOffice2PDF cmd exitStatus {}", exitStatus);
        } else {
            log.debug("convertOffice2PDF cmd exitStatus {}", exitStatus);
        }
        process.destroy(); // 销毁子进程
        log.info("转化结换.......");
        return true;
    }


    private IConverter converter = null;
    public void office2PdfByMSOffice(InputStream inputStream,OutputStream outputStream) throws Exception{
        String tempPath=PathUtils.getStaticPath()+PathUtils.getSeparator()+"temp";
        File baseFolder=new File(tempPath);
        if(!baseFolder.exists()){
            baseFolder.mkdirs();
        }
        try{
            if(converter==null ) {
                converter = LocalConverter.builder().baseFolder(baseFolder).workerPool(20, 25, 5, TimeUnit.SECONDS).processTimeout(5L, TimeUnit.MINUTES).build();
            }
            converter.convert(inputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            log.debug("Convert success.");
        }catch (Exception e){
            log.warn("转换失败---"+e.getMessage());
        }finally {
            outputStream.flush();
            TimeUnit.MILLISECONDS.sleep(200);
            outputStream.close();
        }
    }


    public File saveTempFile(InputStream inputStream,String fileName) throws Exception{
        String tempPath=PathUtils.getStaticPath()+PathUtils.getSeparator()+"temp";
        File baseFolder=new File(tempPath);
        if(!baseFolder.exists()){
            baseFolder.mkdirs();
        }
        File tempFile=new File(tempPath+PathUtils.getSeparator()+fileName);
        try( InputStream _inputStream=inputStream;
             OutputStream tempOutputStream = new FileOutputStream(tempFile)){
            byte[] buf = new byte[1024];
            int len;
            while((len=_inputStream.read(buf))>0){
                tempOutputStream.write(buf,0,len);
            }
            _inputStream.close();
            tempOutputStream.flush();
            tempOutputStream.close();
        }catch (Exception e){
            e.printStackTrace();
        }
        return tempFile;
    }

}
