package com.example.selenium.service;
import jakarta.mail.MessagingException;
import jakarta.mail.internet.MimeMessage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.core.io.FileSystemResource;

import java.text.SimpleDateFormat;

import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import java.io.File;
import java.util.Calendar;
import java.util.Date;

@Service
public class EmailService {

    @Autowired
    private JavaMailSender javaMailSender;
    private static final Logger logger = LoggerFactory.getLogger(EmailService.class);

    public void sendEmail(String[] to, String subject, String text, String attachmentPath) throws MessagingException {
        Calendar NowDate = Calendar.getInstance();
        NowDate.add(Calendar.DAY_OF_MONTH, -1);
        NowDate.set(Calendar.HOUR_OF_DAY, 0);
        NowDate.set(Calendar.MINUTE, 0);
        NowDate.set(Calendar.SECOND, 0);
        SimpleDateFormat DateFormat = new SimpleDateFormat("yyyy-MM-dd");
        Date date = NowDate.getTime();
        String fileDate = DateFormat.format(date);
        String fileName = fileDate + "news.zip";
        MimeMessage message = javaMailSender.createMimeMessage();
        try {
            MimeMessageHelper helper = new MimeMessageHelper(message, true);
            helper.setTo(to);
            helper.setSubject(subject);
            helper.setText(text, true);

            FileSystemResource file = new FileSystemResource(new File(attachmentPath));
            helper.addAttachment(fileName, file);

            javaMailSender.send(message);
            logger.info("Email sent successfully!");
        } catch (MessagingException e) {
            logger.warn("Email sending failed!", e);
        }
    }

    public void sendEmail(String[] to, String subject, String text) throws MessagingException {
        MimeMessage message = javaMailSender.createMimeMessage();
        try {
            MimeMessageHelper helper = new MimeMessageHelper(message, true);
            helper.setTo(to);
            helper.setSubject(subject);
            helper.setText(text, true);
            javaMailSender.send(message);
            logger.info("Error message sent successfully!");
        } catch (MessagingException e) {
            logger.info("Error message sending failed!", e);
        }
    }
}
