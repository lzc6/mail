package com.example.mail.controller;


import com.example.mail.domain.ResultMap;
import com.example.mail.services.BillFile;
import com.example.mail.services.MailService;
import com.example.mail.utils.DateTimeUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.mail.MessagingException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;

@Component
@RestController
@RequestMapping(value = "/mail")
public class MailController {
    @Autowired
    private MailService mailService;
    //收件人
    private String[] test ={"picczq14@picczq.com"};
//    private String[] test ={"liangzhicong@test.com"};

//肇庆
    private String[] zq ={"picczq14@picczq.com","kongyuji@picczq.com","chenjianqiang@picczq.com","ouwanchao@picczq.com",
        "liangzhijian@picczq.com","linyi@picczq.com","chenshishen@picczq.com","liangguoyi@picczq.com","wangwei@picczq.com","luyongjin@picczq.com"};
//    鼎湖
    private String[] dh ={"picczq14@picczq.com","shaozuhui@picczq.com"};
//    广宁
    private String[] gn ={"picczq14@picczq.com","huangjianjun@picczq.com"};
//    四会
    private String[] sh ={"picczq14@picczq.com","liangguoqian@picczq.com"};
//    怀集
    private String[] hj ={"picczq14@picczq.com","chensi@picczq.com","liangweituan@picczq.com"};
//    高要
    private String[] gy ={"picczq14@picczq.com","tanshu@picczq.com"};
//    德庆
    private String[] dq ={"picczq14@picczq.com","huangweiming@picczq.com"};
//    封开
    private String[] fk ={"picczq14@picczq.com","xuyunyong@picczq.com"};
//    端州
    private String[] dz ={"picczq14@picczq.com","zhanglan@picczq.com"};
//    车商
    private String[] cs ={"picczq14@picczq.com","ouyaohao@picczq.com"};
//    黄凌志
    private String[] huanglingzhi ={"huanglingzhi@picczq.com","picczq14@picczq.com"};
//    符嘉敏
    private String[] fujiamin ={"fujiamin@picczq.com","picczq14@picczq.com"};
//    出单中心
    private String[] cdtj ={"konghao@picczq.com","limeizhuang@picczq.com","picczq14@picczq.com"};

//    黄玉芳
    private String[] huangyufang = {"huangyufang@picczq.com","picczq14@picczq.com"};
//    车险部 王伟
    private String[] wangwei = {"wangwei@picczq.com","picczq14@picczq.com"};

    @Autowired
    private BillFile billFile;
    static {
        System.setProperty("mail.mime.splitlongparameters","false");
    }


    int monthminonenum = DateTimeUtils.getmonthminonenum();
    String nextmonth = DateTimeUtils.getmonthaddone();
    String nexttwomonth = DateTimeUtils.getmonthaddtow();
    String lastweek = DateTimeUtils.getlastweek();
    String month = DateTimeUtils.getmonth();
    String todayminone = DateTimeUtils.getdateminone();
    //每天
    @GetMapping("/dailysend")
    public ResultMap dailsendMail(){


//        .//data//
        //报案核损送修结案
        String title1 = todayminone+"肇庆日度清单";
        String title2 = todayminone+"鼎湖日度清单";
        String title3 = todayminone+"广宁日度清单";
        String title4 = todayminone+"四会日度清单";
        String title5 = todayminone+"怀集日度清单";
        String title6 = todayminone+"高要日度清单";
        String title7 = todayminone+"德庆日度清单";
        String title8 = todayminone+"封开日度清单";
        String title9 = todayminone+"端州日度清单";
        String title10 = todayminone+"车商日度清单";
        String content =todayminone+"报案清单、送修清单、核损清单、结案清单";



        String[] zqfile = {"D:\\bill\\"+ todayminone + "-肇庆-每日报案清单.xls","D:\\bill\\" + todayminone + "-肇庆-每日送修清单.xls","D:\\bill\\" + todayminone + "-肇庆-每日核损清单.xls"};
        String[] dhfile = {"D:\\bill\\" + todayminone + "-鼎湖-每日报案清单.xls","D:\\bill\\" + todayminone + "-鼎湖-每日送修清单.xls","D:\\bill\\" + todayminone + "-鼎湖-每日核损清单.xls"};
        String[] gnfile = {"D:\\bill\\" + todayminone + "-广宁-每日报案清单.xls","D:\\bill\\" + todayminone + "-广宁-每日送修清单.xls","D:\\bill\\" + todayminone + "-广宁-每日核损清单.xls"};
        String[] shfile = {"D:\\bill\\" + todayminone + "-四会-每日报案清单.xls","D:\\bill\\" + todayminone + "-四会-每日送修清单.xls","D:\\bill\\" + todayminone + "-四会-每日核损清单.xls"};
        String[] hjfile = {"D:\\bill\\" + todayminone + "-怀集-每日报案清单.xls","D:\\bill\\" + todayminone + "-怀集-每日送修清单.xls","D:\\bill\\" + todayminone + "-怀集-每日核损清单.xls"};
        String[] gyfile = {"D:\\bill\\" + todayminone + "-高要-每日报案清单.xls","D:\\bill\\" + todayminone + "-高要-每日送修清单.xls","D:\\bill\\" + todayminone + "-高要-每日核损清单.xls"};
        String[] dqfile = {"D:\\bill\\" + todayminone + "-德庆-每日报案清单.xls","D:\\bill\\" + todayminone + "-德庆-每日送修清单.xls","D:\\bill\\" + todayminone + "-德庆-每日核损清单.xls"};
        String[] fkfile = {"D:\\bill\\" + todayminone + "-封开-每日报案清单.xls","D:\\bill\\" + todayminone + "-封开-每日送修清单.xls","D:\\bill\\" + todayminone + "-封开-每日核损清单.xls"};
        String[] dzfile = {"D:\\bill\\" + todayminone + "-端州-每日报案清单.xls","D:\\bill\\" + todayminone + "-端州-每日送修清单.xls","D:\\bill\\" + todayminone + "-端州-每日核损清单.xls"};
        String[] csfile = {"D:\\bill\\" + todayminone + "-车商-每日报案清单.xls","D:\\bill\\" + todayminone + "-车商-每日送修清单.xls","D:\\bill\\" + todayminone + "-车商-每日核损清单.xls"};

//        String[] zqfile = {".//data//" + todayminone + "-肇庆-每日报案清单.xls",".//data//" + todayminone + "-肇庆-每日送修清单.xls",".//data//" + todayminone + "-肇庆-每日核损清单.xls",".//data//" + todayminone + "-肇庆-每日结案清单.xls"};
//        String[] dhfile = {".//data//" + todayminone + "-鼎湖-每日报案清单.xls",".//data//" + todayminone + "-鼎湖-每日送修清单.xls",".//data//" + todayminone + "-鼎湖-每日核损清单.xls",".//data//" + todayminone + "-鼎湖-每日结案清单.xls"};
//        String[] gnfile = {".//data//" + todayminone + "-广宁-每日报案清单.xls",".//data//" + todayminone + "-广宁-每日送修清单.xls",".//data//" + todayminone + "-广宁-每日核损清单.xls",".//data//" + todayminone + "-广宁-每日结案清单.xls"};
//        String[] shfile = {".//data//" + todayminone + "-四会-每日报案清单.xls",".//data//" + todayminone + "-四会-每日送修清单.xls",".//data//" + todayminone + "-四会-每日核损清单.xls",".//data//" + todayminone + "-四会-每日结案清单.xls"};
//        String[] hjfile = {".//data//" + todayminone + "-怀集-每日报案清单.xls",".//data//" + todayminone + "-怀集-每日送修清单.xls",".//data//" + todayminone + "-怀集-每日核损清单.xls",".//data//" + todayminone + "-怀集-每日结案清单.xls"};
//        String[] gyfile = {".//data//" + todayminone + "-高要-每日报案清单.xls",".//data//" + todayminone + "-高要-每日送修清单.xls",".//data//" + todayminone + "-高要-每日核损清单.xls",".//data//" + todayminone + "-高要-每日结案清单.xls"};
//        String[] dqfile = {".//data//" + todayminone + "-德庆-每日报案清单.xls",".//data//" + todayminone + "-德庆-每日送修清单.xls",".//data//" + todayminone + "-德庆-每日核损清单.xls",".//data//" + todayminone + "-德庆-每日结案清单.xls"};
//        String[] fkfile = {".//data//" + todayminone + "-封开-每日报案清单.xls",".//data//" + todayminone + "-封开-每日送修清单.xls",".//data//" + todayminone + "-封开-每日核损清单.xls",".//data//" + todayminone + "-封开-每日结案清单.xls"};
//        String[] dzfile = {".//data//" + todayminone + "-端州-每日报案清单.xls",".//data//" + todayminone + "-端州-每日送修清单.xls",".//data//" + todayminone + "-端州-每日核损清单.xls",".//data//" + todayminone + "-端州-每日结案清单.xls"};
//        String[] csfile = {".//data//" + todayminone + "-车商-每日报案清单.xls",".//data//" + todayminone + "-车商-每日送修清单.xls",".//data//" + todayminone + "-车商-每日核损清单.xls",".//data//" + todayminone + "-车商-每日结案清单.xls"};

        ResultMap baoan = billFile.baoan();
        ResultMap songxiu = billFile.songxiu();
        ResultMap hesun = billFile.hesun();
        ResultMap jiean = billFile.jiean();
        //保呗商城
        String title11 = todayminone+"每日保贝商城清单、每日保贝商城积分汇总清单";
        String content11 =todayminone+"每日保贝商城清单、每日保贝商城积分汇总清单";
        String[] bbscfile = {".//data//" + todayminone +"-每日保贝商城清单.xls",".//data//" + todayminone +"-每日保贝商城积分汇总清单.xls"};
        ResultMap bbsc = billFile.bbsc();
        try {
            mailService.sendAttachmentsMail(zq, title1, content, zqfile);
//            mailService.sendAttachmentsMail(dh, title2, content, dhfile);
//            mailService.sendAttachmentsMail(gn, title3, content, gnfile);
//            mailService.sendAttachmentsMail(sh, title4, content, shfile);
//            mailService.sendAttachmentsMail(hj, title5, content, hjfile);
//            mailService.sendAttachmentsMail(gy, title6, content, gyfile);
//            mailService.sendAttachmentsMail(dq, title7, content, dqfile);
//            mailService.sendAttachmentsMail(fk, title8, content, fkfile);
//            mailService.sendAttachmentsMail(dz, title9, content, dzfile);
//            mailService.sendAttachmentsMail(cs, title10, content, csfile);
//            mailService.sendAttachmentsMail(wangwei, title11, content11, bbscfile);
        } catch (MessagingException e) {
            return new ResultMap(0,"",true,e,"");
        } catch (UnsupportedEncodingException e) {
            return new ResultMap(0,"",true,e,"");
        }

        return new ResultMap(1,"",true,"","");
    }
    //每月1号
    @GetMapping("/monthsend")
    public ResultMap monthsendMail() {

//        .//data//
        //交叉销售-黄凌志  车商口径-黄凌志
        //        ResultMap jiaochaxiaoshou = billFile.jiaocha();
//        ResultMap cheshangkoujing = billFile.cheshangkoujing();
//        ResultMap chexianqibao = billFile.chexianqibao();
//        ResultMap chexianqiandan = billFile.chexianqiandan();
//        ResultMap feicheqibao = billFile.feicheqibao();
//        ResultMap feicheqiandan = billFile.feicheqiandan();
        String title1A = monthminonenum+"月销管清单-交叉销售清单";
        String content1A =monthminonenum+"月车险交叉销售清单，非车交叉销售清单";
        String title1B = monthminonenum+"月销管清单-车商口径清单";
        String content1B =monthminonenum+"月销管清单-车商口径清单";
        String title1C = monthminonenum+"月销管清单-车险起保，车险签单清单";
        String content1C =monthminonenum+"月销管清单-车险起保，车险签单清单";
        String title1D = monthminonenum+"月销管清单-非车起保，非车签单清单";
        String content1D =monthminonenum+"月销管清单-非车起保，非车签单清单";
//        String[] xgfile1A = {"D:\\bill\\" + monthminonenum +"月车险交叉销售清单.xls","D:\\bill\\" + monthminonenum + "月非车交叉销售清单.xls"};
//        String[] xgfile1B = {"D:\\bill\\" + monthminonenum +"月车商口径清单.xls"};
//        String[] xgfile1C = {"D:\\bill\\" + monthminonenum +"月车险起保清单.xls","D:\\bill\\" + monthminonenum +"月车险签单清单.xls"};
//        String[] xgfile1D = {"D:\\bill\\" + monthminonenum +"月非车起保清单.xls","D:\\bill\\" + monthminonenum +"月非车签单清单.xls"};

        String[] xgfile1A = {".//data//" + monthminonenum +"月车险交叉销售清单.xls",".//data//" + monthminonenum + "月非车交叉销售清单.xls"};
        String[] xgfile1B = {".//data//" + monthminonenum +"月车商口径清单.xls"};
        String[] xgfile1C = {".//data//" + monthminonenum +"月车险起保清单.xls",".//data//" + monthminonenum +"月车险签单清单.xls"};
        String[] xgfile1D = {".//data//" + monthminonenum +"月非车起保清单.xls",".//data//" + monthminonenum +"月非车签单清单.xls"};


        //车船税-符嘉敏
        String title2 = monthminonenum+"月财务车船税清单";
        String content2 =monthminonenum+"月车船税清单";
//        String[] ccsfile = {"D:\\bill\\" + monthminonenum +"月车船税清单.xls"};

        String[] ccsfile = {".//data//" + monthminonenum +"月车船税清单.xls"};



//
        //出单统计量-孔豪
        String title3 = monthminonenum+"月出单中心出单统计量清单";
        String content3 =monthminonenum+"月出单统计量清单";
//        String[] cdtjlfile = {".//data//" + monthminonenum +"月非车险出单.xls",".//data//" + monthminonenum +"月车商车险出单.xls",".//data//" + monthminonenum +"月三版车险出单.xls",".//data//" + monthminonenum +"月非车商个代车险出单.xls",
//                              ".//data//" + monthminonenum +"月车商保单批单.xls",".//data//" + monthminonenum +"月三版保单批单.xls",".//data//" + monthminonenum +"月非车商个代保单批单.xls"};
        String[] cdtjlfile = {"D:\\bill\\" + monthminonenum +"月车商车险出单.xls","D:\\bill\\" + monthminonenum +"月三版车险出单.xls","D:\\bill\\" + monthminonenum +"月非车商个代车险出单.xls",
                              "D:\\bill\\" + monthminonenum +"月车商保单批单.xls","D:\\bill\\" + monthminonenum +"月三版保单批单.xls","D:\\bill\\" + monthminonenum +"月非车商个代保单批单.xls"};



        //报案 送修 核损 结案清单

        String title4A = monthminonenum+"月肇庆月度清单";
        String title4B = monthminonenum+"月鼎湖月度清单";
        String title4C = monthminonenum+"月广宁月度清单";
        String title4D = monthminonenum+"月四会月度清单";
        String title4E = monthminonenum+"月怀集月度清单";
        String title4F = monthminonenum+"月高要月度清单";
        String title4G = monthminonenum+"月德庆月度清单";
        String title4H = monthminonenum+"月封开月度清单";
        String title4I = monthminonenum+"月端州月度清单";
        String title4J = monthminonenum+"月车商月度清单";
        String content4 ="报案清单、送修清单、核损清单、结案清单";
//        String[] zqfile = {"D:\\bill\\肇庆-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\肇庆-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\肇庆-" + monthminonenum + "月核损清单.xls","D:\\bill\\肇庆-" + monthminonenum + "月结案清单.xls"};
//        String[] dhfile = {"D:\\bill\\鼎湖-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\鼎湖-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\鼎湖-" + monthminonenum + "月核损清单.xls","D:\\bill\\鼎湖-" + monthminonenum + "月结案清单.xls"};
//        String[] gnfile = {"D:\\bill\\广宁-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\广宁-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\广宁-" + monthminonenum + "月核损清单.xls","D:\\bill\\广宁-" + monthminonenum + "月结案清单.xls"};
//        String[] shfile = {"D:\\bill\\四会-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\四会-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\四会-" + monthminonenum + "月核损清单.xls","D:\\bill\\四会-" + monthminonenum + "月结案清单.xls"};
//        String[] hjfile = {"D:\\bill\\怀集-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\怀集-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\怀集-" + monthminonenum + "月核损清单.xls","D:\\bill\\怀集-" + monthminonenum + "月结案清单.xls"};
//        String[] gyfile = {"D:\\bill\\高要-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\高要-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\高要-" + monthminonenum + "月核损清单.xls","D:\\bill\\高要-" + monthminonenum + "月结案清单.xls"};
//        String[] dqfile = {"D:\\bill\\德庆-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\德庆-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\德庆-" + monthminonenum + "月核损清单.xls","D:\\bill\\德庆-" + monthminonenum + "月结案清单.xls"};
//        String[] fkfile = {"D:\\bill\\封开-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\封开-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\封开-" + monthminonenum + "月核损清单.xls","D:\\bill\\封开-" + monthminonenum + "月结案清单.xls"};
//        String[] dzfile = {"D:\\bill\\端州-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\端州-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\端州-" + monthminonenum + "月核损清单.xls","D:\\bill\\端州-" + monthminonenum + "月结案清单.xls"};
//        String[] csfile = {"D:\\bill\\车商-"+  monthminonenum+  "月报案清单.xls","D:\\bill\\车商-" + monthminonenum + "月送修清单.xls",
//                           "D:\\bill\\车商-" + monthminonenum + "月核损清单.xls","D:\\bill\\车商-" + monthminonenum + "月结案清单.xls"};


        String[] zqfile = {".//data//肇庆-"+  monthminonenum+  "月报案清单.xls",".//data//肇庆-" + monthminonenum + "月送修清单.xls",
                ".//data//肇庆-" + monthminonenum + "月核损清单.xls",".//data//肇庆-" + monthminonenum + "月结案清单.xls"};
        String[] dhfile = {".//data//鼎湖-"+  monthminonenum+  "月报案清单.xls",".//data//鼎湖-" + monthminonenum + "月送修清单.xls",
                ".//data//鼎湖-" + monthminonenum + "月核损清单.xls",".//data//鼎湖-" + monthminonenum + "月结案清单.xls"};
        String[] gnfile = {".//data//广宁-"+  monthminonenum+  "月报案清单.xls",".//data//广宁-" + monthminonenum + "月送修清单.xls",
                ".//data//广宁-" + monthminonenum + "月核损清单.xls",".//data//广宁-" + monthminonenum + "月结案清单.xls"};
        String[] shfile = {".//data//四会-"+  monthminonenum+  "月报案清单.xls",".//data//四会-" + monthminonenum + "月送修清单.xls",
                ".//data//四会-" + monthminonenum + "月核损清单.xls",".//data//四会-" + monthminonenum + "月结案清单.xls"};
        String[] hjfile = {".//data//怀集-"+  monthminonenum+  "月报案清单.xls",".//data//怀集-" + monthminonenum + "月送修清单.xls",
                ".//data//怀集-" + monthminonenum + "月核损清单.xls",".//data//怀集-" + monthminonenum + "月结案清单.xls"};
        String[] gyfile = {".//data//高要-"+  monthminonenum+  "月报案清单.xls",".//data//高要-" + monthminonenum + "月送修清单.xls",
                ".//data//高要-" + monthminonenum + "月核损清单.xls",".//data//高要-" + monthminonenum + "月结案清单.xls"};
        String[] dqfile = {".//data//德庆-"+  monthminonenum+  "月报案清单.xls",".//data//德庆-" + monthminonenum + "月送修清单.xls",
                ".//data//德庆-" + monthminonenum + "月核损清单.xls",".//data//德庆-" + monthminonenum + "月结案清单.xls"};
        String[] fkfile = {".//data//封开-"+  monthminonenum+  "月报案清单.xls",".//data//封开-" + monthminonenum + "月送修清单.xls",
                ".//data//封开-" + monthminonenum + "月核损清单.xls",".//data//封开-" + monthminonenum + "月结案清单.xls"};
        String[] dzfile = {".//data//端州-"+  monthminonenum+  "月报案清单.xls",".//data//端州-" + monthminonenum + "月送修清单.xls",
                ".//data//端州-" + monthminonenum + "月核损清单.xls",".//data//端州-" + monthminonenum + "月结案清单.xls"};
        String[] csfile = {".//data//车商-"+  monthminonenum+  "月报案清单.xls",".//data//车商-" + monthminonenum + "月送修清单.xls",
                ".//data//车商-" + monthminonenum + "月核损清单.xls",".//data//车商-" + monthminonenum + "月结案清单.xls"};

        //保呗商城
        String title5 = monthminonenum+"月保贝商城清单、保贝商城积分汇总清单";
        String content5 =monthminonenum+"月保贝商城清单、保贝商城积分汇总清单";
        String[] bbscfile = {".//data//" + monthminonenum +"月保贝商城清单.xls",".//data//" + monthminonenum +"月保贝商城积分汇总清单.xls"};

        //异地车
        String title6 = monthminonenum+"月异地车清单";
        String content6 = monthminonenum+"月异地车清单";
//        String[] ydcfile = {"\\\\H:\\bill\\" + monthminonenum +"-"+"月异地车清单.xls"};
        String[] ydcfile = {".//data//" + monthminonenum +"月异地车清单.xls"};

        //生成excle
        //异地车
        ResultMap ydc = billFile.ydc();
        //保呗商城
        ResultMap bbsc = billFile.bbscmonth();
        //报案 送修 核损 结案清单
        ResultMap baoan = billFile.monthbaoan();
        ResultMap songxiu = billFile.monthsongxiu();
        ResultMap hesun = billFile.monthhesun();
        ResultMap jiean = billFile.monthjiean();
        //出单统计量-孔豪
        ResultMap cdtjl = billFile.cdtjl();
        //车船税-符嘉敏
        ResultMap chechuanshui = billFile.chechuanshui();
        //交叉销售-黄凌志  车商口径-黄凌志
        ResultMap jiaochaxiaoshou = billFile.jiaocha();
        ResultMap cheshangkoujing = billFile.cheshangkoujing();
        ResultMap chexianqibao = billFile.chexianqibao();
        ResultMap chexianqiandan = billFile.chexianqiandan();
        ResultMap feicheqibao = billFile.feicheqibao();
        ResultMap feicheqiandan = billFile.feicheqiandan();


        try {
            mailService.sendAttachmentsMail(huanglingzhi, title1A, content1A, xgfile1A);
            mailService.sendAttachmentsMail(huanglingzhi, title1B, content1B, xgfile1B);
            mailService.sendAttachmentsMail(huanglingzhi, title1C, content1C, xgfile1C);
            mailService.sendAttachmentsMail(huanglingzhi, title1D, content1D, xgfile1D);
            mailService.sendAttachmentsMail(fujiamin, title2, content2, ccsfile);
            mailService.sendAttachmentsMail(cdtj, title3, content3, cdtjlfile);
            mailService.sendAttachmentsMail(zq, title4A, content4, zqfile);
            mailService.sendAttachmentsMail(dh, title4B, content4, dhfile);
            mailService.sendAttachmentsMail(gn, title4C, content4, gnfile);
            mailService.sendAttachmentsMail(sh, title4D, content4, shfile);
            mailService.sendAttachmentsMail(hj, title4E, content4, hjfile);
            mailService.sendAttachmentsMail(gy, title4F, content4, gyfile);
            mailService.sendAttachmentsMail(dq, title4G, content4, dqfile);
            mailService.sendAttachmentsMail(fk, title4H, content4, fkfile);
            mailService.sendAttachmentsMail(dz, title4I, content4, dzfile);
            mailService.sendAttachmentsMail(cs, title4J, content4, csfile);
            mailService.sendAttachmentsMail(wangwei, title5, content5, bbscfile);
            mailService.sendAttachmentsMail(wangwei, title6, content6, ydcfile);
        } catch (MessagingException e) {
            return new ResultMap(0,"",true,e,"");
        } catch (UnsupportedEncodingException e) {
            return new ResultMap(0,"",true,e,"");
        }

        return new ResultMap(1,"",true,"","");
    }
    //每月10日
    @GetMapping("/midmonthsend")
    public ResultMap midmonthsendMail() throws IOException {

        //邮政业务
        String title1 = nextmonth+"-"+nexttwomonth+"非车险邮政业务到期清单";
        String content1 =nextmonth+"-"+nexttwomonth+"非车险邮政业务到期清单";
//        String[] yzdqfile = {"D:\\bill\\" + nextmonth+"-"+nexttwomonth +"邮政到期清单.xls","D:\\bill\\" + nextmonth+"-"+nexttwomonth +"到期电话号码.txt"};
        String[] yzdqfile = {".//data//" + nextmonth+"-"+nexttwomonth +"邮政到期清单.xls",".//data//" + nextmonth+"-"+nexttwomonth +"到期电话号码.txt"};

        ResultMap youzhengdaoqi = billFile.yzdq();

        try {
              mailService.sendAttachmentsMail(huangyufang, title1, content1, yzdqfile);

        } catch (MessagingException e) {
            return new ResultMap(0,"",true,e,"");
        } catch (UnsupportedEncodingException e) {
            return new ResultMap(0,"",true,e,"");
        }

        return new ResultMap(1,"",true,"","");

    }
    //每周
    @GetMapping("/weeksend")
    public ResultMap weeksend(){




        //批改清单-出单中心
        String title1 = lastweek+"-"+todayminone+"批改清单";
        String content =lastweek+"-"+todayminone+"批改清单";
        ResultMap pigai = billFile.pigai();
        String[] cdzxfile = {".//data//"+ lastweek+"-"+todayminone+"批改清单.xls"};
//        String[] cdzxfile = {"D:\\bill\\"+ lastweek+"-"+todayminone+"批改清单.xls"};

        //销管-车险起保-非车起保
//        String title2 = month+"-"+todayminone+"车险起保清单、非车起保清单";
//        String content2 =month+"-"+todayminone+"车险起保清单、非车起保清单";
//
//        ResultMap chexianqibao = billFile.chexianqibaoweek();
//        ResultMap feicheqibao = billFile.feicheqibaoweek();
        try {
            mailService.sendAttachmentsMail(cdtj, title1, content, cdzxfile);
        } catch (MessagingException e) {
            return new ResultMap(0,"",true,e,"");
        } catch (UnsupportedEncodingException e) {
            return new ResultMap(0,"",true,e,"");
        }

        return new ResultMap(1,"",true,"","");
    }


}



