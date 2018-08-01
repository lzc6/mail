package com.example.mail.services;

import com.example.mail.domain.ResultMap;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeUtility;
import java.io.File;
import java.io.UnsupportedEncodingException;

@Service
public class MailService {

    @Value("${spring.mail.username}")
    private String from;

    @Autowired
    private JavaMailSender sender;

    /*发送邮件的方法*/
    //纯文本
//    public ResultMap sendSimple(String to, String title, String content){
//        SimpleMailMessage message = new SimpleMailMessage();
//        message.setFrom(from); //发送者
//        message.setTo(to); //接受者
//        message.setSubject(title); //发送标题
//        message.setText(content); //发送内容
//        try {
//            sender.send(message);
//            return new ResultMap(0, "发送成功", true, "", "");
//        } catch (Exception e) {
//            return new ResultMap(0, "发送失败", true, e, "");
//        }
    //带附件
//    String filePath1,String filePath2,String filePath3
    public ResultMap sendAttachmentsMail(String[] to, String subject, String content,String[] filepath) throws MessagingException, UnsupportedEncodingException {
        MimeMessage message = sender.createMimeMessage();
        MimeMessageHelper helper = new MimeMessageHelper(message, true);
        helper.setFrom(from);
        helper.setTo(to);
        helper.setSubject(subject);
        helper.setText(content, true);

        for (int i = 0;i<filepath.length;i++){
            FileSystemResource file = new FileSystemResource(new File(filepath[i]));
            String filename = filepath[i].substring(filepath[i].lastIndexOf(File.separator)+1);

            helper.addAttachment(filename, file);

        }

        try{
            sender.send(message);
            return new ResultMap(0, "发送成功", true, "", "");
        }catch (Exception e){
            return new ResultMap(0, "发送失败", true, e, "");
        }
    }
}