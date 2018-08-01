package com.example.mail.services;
import com.example.mail.domain.ResultMap;
import com.example.mail.utils.DateTimeUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.web.bind.annotation.RequestMapping;
import com.example.mail.utils.TxtExport;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Service
public class BillFile {
    String today = DateTimeUtils.getdatetime();
    String todayminone = DateTimeUtils.getdateminone();
    String lastmonth = DateTimeUtils.getmonthminone();
    String month = DateTimeUtils.getmonth();
    String lastyearmonth = DateTimeUtils.getlastyearmonth();
    String lastyearlastmonth = DateTimeUtils.getlastyearlastmonth();
    String nextmonth = DateTimeUtils.getmonthaddone();
    String nexttowmonth = DateTimeUtils.getmonthaddtow();
    String lastweek = DateTimeUtils.getlastweek();
    @Autowired
    private ExportToExcel exportToExcel;

    @Autowired
    @Qualifier("picczq")
    protected JdbcTemplate picczq;
    //报案
    public ResultMap baoan(){
        picczq.update("create temporary table baoan\n" +
                "select  a.makecom,c.licenseno,c.policyno,a.registno,a.reportdate,a.reportorname,a.reportornumber,a.damagedate,a.damagehour,a.damageaddress,a.damagename,\n" +
                "b.lflag,b.checker1,truncate(b.sumestimatefee,2),b.firstsiteflag \n" +
                "from cl_prplregist a,\n" +
                "cl_prplregistsummary c,\n" +
                "cl_prplchecktask b\n" +
                "where a.validflag='1'\n" +
                "and a.reportdate >= ? \n" +
                "and a.reportdate < ? \n" +
                "and a.registno=c.registno \n" +
                "and a.registno=b.registno;",todayminone,today);
        picczq.update("create index policyno on baoan(policyno);");
        picczq.update("create temporary table baoan1\n" +
                "select a.*,b.handlercode,cast('' as nchar(100)) as handlername,b.handler1code,\n" +
                "cast('' as nchar(100)) as handler1name\n" +
                "from baoan a,\n" +
                "cx_prpcmain b\n" +
                "where a.policyno=b.policyno");
        picczq.update("update baoan1 set handlername = (select max(username) from utiihandler a \n" +
                "where baoan1.handlercode = a.usercode);");
        picczq.update("update baoan1 set handler1name = (select max(username) from utiihandler a \n" +
                "where baoan1.handler1code = a.usercode);");

        //肇庆
        List<Map<String, Object>> result1 = picczq.queryForList("select * from baoan1 order by reportdate desc;");
        //车商

        List<Map<String, Object>> result2 = picczq.queryForList("select * from baoan1\n" +
                "where makecom in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342') \n" +
                "order by reportdate desc;");
        //端州
        List<Map<String, Object>> result3 = picczq.queryForList("select * from baoan1\n" +
                "where makecom in ('44129311','44120100','44120313','44129301','44129302','44129303','44129315','44129316','44129317','44129318','44129319','44129321','44129324','44129329','44129330','44129338')\n" +
                "order by reportdate desc;");
        //鼎湖
        List<Map<String, Object>> result4 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like '441203%' and makecom <>'44120313'\n" +
                "order by reportdate desc;");
        //四会

        List<Map<String, Object>> result5 = picczq.queryForList("select * from baoan1\n" +
                "where makecom  like  '441284%'" +
                "or makecom in ('44129323','44129325')" +
                "order by reportdate desc;");
        //广宁

        List<Map<String, Object>> result6 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441223%'\n" +
                "order by reportdate desc;");
        //怀集

        List<Map<String, Object>> result7 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441224%'\n" +
                "order by reportdate desc;");
        //高要

        List<Map<String, Object>> result8 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441283%'\n" +
                "order by reportdate desc;");
        //德庆

        List<Map<String, Object>> result9 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441226%'\n" +
                "order by reportdate desc;");
        //封开

        List<Map<String, Object>> result10 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441225%'\n" +
                "order by reportdate desc;");

        picczq.update("drop table baoan;");
        picczq.update("drop table baoan1;");

        exportToExcel.baoan1(result1,"肇庆");


        //车商
        exportToExcel.baoan1(result2,"车商");

        //端州
        exportToExcel.baoan1(result3,"端州");

        //鼎湖
        exportToExcel.baoan1(result4,"鼎湖");

        //四会
        exportToExcel.baoan1(result5,"四会");

        //广宁
        exportToExcel.baoan1(result6,"广宁");

        //怀集
        exportToExcel.baoan1(result7,"怀集");

        //高要
        exportToExcel.baoan1(result8,"高要");

        //德庆
        exportToExcel.baoan1(result9,"德庆");

        //封开
        exportToExcel.baoan1(result10,"封开");




        return null;
    }
    //送修
    public ResultMap songxiu(){
        picczq.update("create temporary table songxiu\n" +
                "select a2.comcode,a0.registno,b0.claimno,a2.lflag,b1.licenseno,b1.brandname,b0.policyno,b0.insuredname,b0.agentcode,b0.startdate,\n" +
                "b0.enddate,b0.businessnature,a0.damagedate,a0.damageaddress,a0.reportdate,a0.reportornumber,a2.repairfactorycode,a2.repairfactoryname,\n" +
                "a2.handlername,a2.deflossdate,a2.underwritename,truncate(a2.sumverilossfee,2),truncate(a2.sumlossfee,2),\n" +
                "cast('' as nchar(100)) as checker1,b.handlercode,\n" +
                "cast('' as nchar(100)) as handname,b.handler1code,\n" +
                "cast('' as nchar(100)) as hand1name,\n" +
                "cast('' as nchar(100)) as agentname,a2.underwriteenddate\n" +
                "from  cl_prplregist a0,cl_prpldeflossmain a2,cl_prplcmain b0,cl_prplcitemcar b1,cx_prpcmain b\n" +
                "where substring(a2.underwriteflag,1,1) in ('1','3') \n" +
                "and a2.validflag='1' \n" +
                "and a2.underwriteenddate >= ?\n" +
                "and a2.underwriteenddate < ?\n" +
                "and a2.registno=b0.registno \n" +
                "and a2.riskcode=b0.riskcode \n" +
                "and b0.id=b1.prplcmainid \n" +
                "and a2.registno=a0.registno\n" +
                "and b0.policyno=b.policyno;",todayminone,today);
        picczq.update("update songxiu set handname = (select max(username) from utiihandler \n" +
                "where songxiu.handlercode = utiihandler.usercode);");
        picczq.update("update songxiu set checker1 = (select max(checker1) from cl_prplchecktask \n" +
                "where songxiu.registno = cl_prplchecktask.registno);");
        picczq.update("update songxiu set hand1name = (select max(username) from utiihandler \n" +
                "where songxiu.handler1code = usercode);");
        picczq.update("update songxiu set agentname = (select agentname from prpdagent \n" +
                "where songxiu.agentcode = prpdagent.agentcode);");

        //肇庆

        List<Map<String, Object>> result1 = picczq.queryForList("select * from songxiu order by UNDERWRITEENDDATE desc;");
        //车商


        List<Map<String, Object>> result2 = picczq.queryForList("select * from songxiu\n" +
                "where comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342') \n" +
                "order by UNDERWRITEENDDATE desc;");
        //端州
        List<Map<String, Object>> result3 = picczq.queryForList("select * from songxiu\n" +
                "where comcode in ('44129311','44120100','44120313','44129301','44129302','44129303','44129315','44129316','44129317','44129318','44129319','44129321','44129324','44129329','44129330','44129338')\n" +
                "order by UNDERWRITEENDDATE desc;");
        //鼎湖

        List<Map<String, Object>> result4 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like '441203%' and comcode <>'44120313'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //四会


        List<Map<String, Object>> result5 = picczq.queryForList("select * from songxiu\n" +
                "where comcode  like  '441284%'" +
                "or comcode in ('44129323','44129325')" +
                "order by UNDERWRITEENDDATE desc;");
        //广宁

        List<Map<String, Object>> result6 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441223%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //怀集
        List<Map<String, Object>> result7 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441224%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //高要

        List<Map<String, Object>> result8 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441283%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //德庆


        List<Map<String, Object>> result9 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441226%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //封开

        List<Map<String, Object>> result10 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441225%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        picczq.update("drop table songxiu");
        //肇庆
        exportToExcel.songxiu1(result1,"肇庆");
        //车商
        exportToExcel.songxiu1(result2,"车商");
        //端州
        exportToExcel.songxiu1(result3,"端州");
        //鼎湖
        exportToExcel.songxiu1(result4,"鼎湖");
        //四会
        exportToExcel.songxiu1(result5,"四会");
        //广宁
        exportToExcel.songxiu1(result6,"广宁");

        //怀集
        exportToExcel.songxiu1(result7,"怀集");
        //高要
        exportToExcel.songxiu1(result8,"高要");;
        //德庆
        exportToExcel.songxiu1(result9,"德庆");
        //封开
        exportToExcel.songxiu1(result10,"封开");

        return null;
    }
    //核损
    public ResultMap hesun(){
        picczq.update("create temporary table temp011_tmp \n" +
                "select a.sumlossfee,a.sumverilossfee,a.cetainlosstype,a.repairfactorytype,a.registno,a.policyno,a.inputtime,a.enddeflossdate, a.finalhandlername,a.underwritecode,\n" +
                "a.underwritename,a.underwriteenddate,a.sumverichgcompfee,a.sumverirepairfee,\n" +
                "(case when left(a.underwriteflag,1) in ('1','3') \n" +
                "then '已核损' else '未核损' end) as  flag1,a.examfactorycode,a.examfactoryname,a.repairfactorycode,a.repairfactoryname,a.prplthirdpartid,0 as reclaim, d.handlercode as reclaimcode, \n" +
                "d.handlername as reclaimname, d.inputtime as reclaiminputtime,d.inputtime as reclaimoutputtime,d.id as reclaimid, a.id as lossmainid, '本地保单' as lflag\n" +
                "from  cl_prpldeflossmain a \n" +
                "inner join cl_prplregistsummary b\n" +
                "on  a.registno = b.registno\n" +
                "and a.riskcode = b.riskcode\n" +
                "and a.enddeflossdate >=?\n" +
                "and a.enddeflossdate <?\n" +
                "and b.comcode like '4412%' \n" +
                "left join  cl_prplreclaim d \n" +
                "on  a.registno = d.registno\n" +
                "and a.id = d.lossmainid\n" +
                "and d.validflag = '1';",todayminone,today);
        picczq.update("insert into temp011_tmp\n" +
                "select  a.sumlossfee,a.sumverilossfee,a.cetainlosstype,a.repairfactorytype,a.registno,a.policyno,a.inputtime,a.enddeflossdate,a.finalhandlername,a.underwritecode,\n" +
                "a.underwritename,a.underwriteenddate,a.sumverichgcompfee,a.sumverirepairfee,(case when left(a.underwriteflag,1)  in ('1','3') \n" +
                "then '已核损' else '未核损' end) as  flag1,a.examfactorycode,a.examfactoryname,a.repairfactorycode,a.repairfactoryname,a.prplthirdpartid,0 as reclaim, d.handlercode as reclaimcode, \n" +
                "d.handlername as reclaimname, d.inputtime as reclaiminputtime, d.inputtime as reclaimoutputtime, \n" +
                "d.id as reclaimid, a.id as ossmainid, '外地保单' as lflag\n" +
                "from  cl_prpldeflossmain a \n" +
                "inner join cl_prplregistsummary b\n" +
                "on a.registno = b.registno\n" +
                "and a.riskcode = b.riskcode\n" +
                "and a.enddeflossdate >=?\n" +
                "and (b.comcode not like '4412%' and a.makecom like  '4412%')                \n" +
                "left join  cl_prplreclaim d \n" +
                "on  a.registno = d.registno\n" +
                "and a.id = d.lossmainid\n" +
                "and d.validflag = '1';",todayminone);
        picczq.update("create temporary table temp011 \t\n" +
                "select  a.sumlossfee,a.sumverilossfee,a.cetainlosstype,a.repairfactorytype,\n" +
                "cast('' as char(20)) as repairfactorytypename,\n" +
                "a.registno,a.lflag,a.policyno,\n" +
                "a.inputtime,a.enddeflossdate, a.finalhandlername,a.underwritecode,a.underwritename,a.underwriteenddate,a.sumverichgcompfee,a.sumverirepairfee,a.flag1,a.examfactorycode,a.examfactoryname,\n" +
                "a.repairfactorycode,a.repairfactoryname,a.prplthirdpartid,reclaim, reclaimcode,reclaimname, reclaiminputtime, reclaimoutputtime, reclaimid, lossmainid\n" +
                "from temp011_tmp a;");
        picczq.update("update temp011  set repairfactorytypename=(select codecname from prpdcode where codetype = 'RepairFactoryType'\n" +
                "and  validstatus = '1'\n" +
                "and codecode=temp011.repairfactorytype);");
        picczq.update("create index inx_3a01a_01 on temp011(prplthirdpartid);");
        picczq.update("create index inx_3a01a_011 on temp011(lossmainid);");
        picczq.update("create index registno on temp011(registno);");
        picczq.update("create index reclaimid on temp011(reclaimid);");
        picczq.update("update temp011\n" +
                "set reclaiminputtime = ((select max(indate)from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.reclaimid\n" +
                "and a.nodeid = '25')),\n" +
                "reclaimoutputtime = ((select max(outdate)from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.reclaimid\n" +
                "and a.nodeid = '25'));");
        picczq.update("update temp011\n" +
                "set reclaimcode= ((select max(a.usercode) from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.lossmainid\n" +
                "and a.nodeid = '25'\n" +
                "and a.valid = '1')),\n" +
                "reclaiminputtime= ((select max(indate)\n" +
                "from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.lossmainid\n" +
                "and a.nodeid = '25'\n" +
                "and a.valid = '1')),\n" +
                "reclaimoutputtime = ((select  max(outdate)\n" +
                "from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.lossmainid\n" +
                "and a.nodeid = '25'\n" +
                "and a.valid = '1'))\n" +
                "where reclaimcode is null;");
        picczq.update("update temp011\n" +
                "set reclaimname = (select username from prpduser a\n" +
                "where  a.usercode = temp011.reclaimcode);");
        picczq.update("update temp011\n" +
                "set reclaim = 1\n" +
                "where reclaimcode is not null;");
        picczq.update("create index inx_3a01a_02 on temp011(registno)");
        picczq.update("create temporary table temp01\n" +
                "select a.*, b.id ,0 as count1,0 as count2,b.brandname,cast('' as nchar(20)) as flag3,cast('非人伤' as nchar(20)) as flag4,cast('' as nchar(20)) as flag5,\n" +
                "cast('' as nchar(20)) as opname1,\n" +
                "cast('' as nchar(20)) as opname2, \n" +
                "cast('' as nchar(20)) as opname3, \n" +
                "cast('' as nchar(20)) as opname4,\n" +
                "cast('1900-01-01 00:00:00' as datetime) as outdate1, \n" +
                "cast('1900-01-01 00:00:00' as datetime) as outdate2, \n" +
                "cast('1900-01-01' as date) as endcasedate\n" +
                "from temp011 a  left join cl_prpldeflossthirdparty b\n" +
                "on a.registno = b.registno \n" +
                "and a.prplthirdpartid = b.id;");
        picczq.update("create index inx_3a01a_03 on temp01(registno);");
        picczq.update("delete from temp01\n" +
                "where exists (select * from  cl_prplregist reg\n" +
                "where reg.registno = temp01.registno\n" +
                "and reg.cancelflag = '1');");
        picczq.update("create temporary table temp02 \n" +
                "select * from  prpduser\n" +
                "where comcode like '4412%';");
        picczq.update("create index inx_tmp02_01 on temp02(usercode)");
        picczq.update("update temp01 \n" +
                "set opname1 = ifnull((select max(b.username) from  cl_prplcompensate a,temp02 b \n" +
                "where a.registno = temp01.registno\n" +
                "and a.operatorcode = b.usercode\n" +
                "and (a.underwriteflag is null or left(a.underwriteflag,1)  not in('7','8'))), '        ')\n" +
                "where flag1 = '已核损';");
        picczq.update("update temp01\n" +
                "set opname2 = (select max(underwritename) from  cl_prplcompensate a\n" +
                "where a.registno = temp01.registno\n" +
                "and (left(a.underwriteflag,1)  in ('1','3')))\n" +
                "where opname1 is not null;");
        picczq.update("update temp01\n" +
                "set flag3 = '已理算未核赔'\n" +
                "where 0 < (select count(*) from  cl_prplcompensate a\n" +
                "where a.registno = temp01.registno\n" +
                "and (a.underwriteflag is null or left(a.underwriteflag,1) in ('0','2','9')));");
        picczq.update("update temp01\n" +
                "set flag3 = '已理算已核赔'\n" +
                "where 0 < (select count(*) from  cl_prplcompensate a\n" +
                "where a.registno = temp01.registno\n" +
                "and (left(a.underwriteflag,1) in ('1','3')))\n" +
                "and flag3 <> '已理算未核赔';");
        picczq.update("update temp01\n" +
                "set flag3 = '已核损未理算'\n" +
                "where flag1= '已核损' and flag3 not like '已理算%';");
        picczq.update("update temp01\n" +
                "set flag3 = '已定损未核损'\n" +
                "where flag1= '未核损';");
        picczq.update("update temp01\n" +
                "set count1 = (select count(*) from cl_prplcomponent b\n" +
                "where b.prpldeflossmainid = temp01.id\n" +
                "and b.registno = temp01.registno\n" +
                "and b.validflag = '1')\n" +
                "where 1 = 1;");
        picczq.update("update temp01\n" +
                "set count2 = (select count(*) from cl_prplrepairfee b\n" +
                "where b.prpldeflossmainid = temp01.id\n" +
                "and b.registno = temp01.registno\n" +
                "and b.validflag = '1');");
        picczq.update("create index  registno on temp01(registno);");
        picczq.update("update temp01\n" +
                "set flag4 = '人伤案'\n" +
                "where 0 < (select count(*) from  cl_prplbpmmain a\n" +
                "where a.mainno = temp01.registno\n" +
                "and a.nodeid = 10\n" +
                "and a.cancelstate='0'\n" +
                "and a.valid ='1');");
        picczq.update("update temp01\n" +
                "set flag5 = ifnull((select max(a.damagecasename)\n" +
                "from  cl_prplcheck a\n" +
                "where a.registno = temp01.registno),' ');");
        picczq.update("create temporary table temp04 \n" +
                "select a.registno, b.usercode, c.username \n" +
                "from temp01 a inner join  cl_prplbpmmain b\n" +
                "on  a.registno = b.mainno\n" +
                "and  b.nodeid = '14'\n" +
                "and b.valid = 1\n" +
                "left join  temp02 c\n" +
                "on b.usercode = c.usercode;");
        picczq.update("create index registno on temp04(registno);");
        picczq.update("update temp01 \n" +
                "set opname3 = ifnull((select max(a.username)\n" +
                "from temp04 a\n" +
                "where a.registno = temp01.registno),'');");
        picczq.update("create temporary table temp05 \n" +
                "select a.registno, b.usercode, b.outdate, c.username          \n" +
                "from temp01 a inner join cl_prplbpmmain b\n" +
                "on a.registno = b.mainno\n" +
                "and b.nodeid = 2\n" +
                "and b.prepnodeid = 4\n" +
                "and b.valid = 1\n" +
                "left join  temp02 c\n" +
                "on b.usercode = c.usercode\n" +
                "order by b.outdate;");
        picczq.update("create temporary table temp055 \n" +
                "select a.registno, b.usercode, b.outdate, c.username          \n" +
                "from temp01 a inner join cl_prplbpmmain b\n" +
                "on a.registno = b.mainno\n" +
                "and b.nodeid = 2\n" +
                "and b.prepnodeid = 4\n" +
                "and b.valid = 1\n" +
                "left join  temp02 c\n" +
                "on b.usercode = c.usercode\n" +
                "order by b.outdate;");
        picczq.update("create temporary table temp06 \n" +
                "select a1.* from temp05 a1, temp055 a2\n" +
                "where a1.registno = a2.registno\n" +
                "and a1.outdate >= a2.outdate\n" +
                "group by a1.registno, a1.usercode, a1.outdate, a1.username\n" +
                "having count(*) <= 1;");
        picczq.update("create index registno on temp06(registno);");
        picczq.update("update temp01 set opname4=ifnull\n" +
                "((select a.username from temp06 a\n" +
                "where a.registno = temp01.registno),'');");
        picczq.update("update temp01 set outdate1=ifnull((select  a.outdate  from temp06 a\n" +
                "where a.registno = temp01.registno),'1900-01-01 00:00:00');");
        picczq.update("update temp01 set outdate1 = null\n" +
                "where outdate1 = '1900-01-01 00:00:00';");
        picczq.update("create temporary table temp07 \n" +
                "select a.registno, b.outdate \n" +
                "from temp01 a,  cl_prplbpmmain b\n" +
                "where a.registno = b.mainno\n" +
                "and b.nodeid = 16\n" +
                "and b.valid = 1\n" +
                "and left(b.businesstype,1) = 'b'\n" +
                "order by b.outdate desc;");
        picczq.update("create temporary table temp077 \n" +
                "select a.registno, b.outdate \n" +
                "from temp01 a,  cl_prplbpmmain b\n" +
                "where a.registno = b.mainno\n" +
                "and b.nodeid = 16\n" +
                "and b.valid = 1\n" +
                "and left(b.businesstype,1) = 'b'\n" +
                "order by b.outdate desc;");
        picczq.update("create index registno on temp07(registno);");
        picczq.update("create index registno on temp077(registno);");
        picczq.update("create temporary table temp08 \n" +
                "select a1.* from temp07 a1, temp077 a2\n" +
                "where a1.registno = a2.registno\n" +
                "and a1.outdate <= a2.outdate\n" +
                "group by a1.registno, a1.outdate\n" +
                "having count(*) <= 1;");
        picczq.update("create index registno on temp08(registno);");
        picczq.update("update temp01 set outdate2 = ifnull((select a.outdate from temp08 a\n" +
                "where a.registno = temp01.registno),'1900-01-01 00:00:00');");
        picczq.update("update temp01 set outdate2 = null\n" +
                "where outdate2 = '1900-01-01 00:00:00';");
        picczq.update("update temp01\n" +
                "set policyno = (select max(policyno) from  cl_prplregistsummary a\n" +
                "where a.registno = temp01.registno)\n" +
                "where policyno is null;");
        picczq.update("create temporary table temp09 \n" +
                "select a.registno, b.endcasedate\n" +
                "from temp01 a,  cl_prplclaim b\n" +
                "where a.registno = b.registno\n" +
                "order by endcasedate desc;");
        picczq.update("create temporary table temp099 \n" +
                "select a.registno, b.endcasedate\n" +
                "from temp01 a,  cl_prplclaim b\n" +
                "where a.registno = b.registno\n" +
                "order by endcasedate desc;");
        picczq.update("create index registno on temp09(registno);");
        picczq.update("create index registno on temp099(registno);");
        picczq.update("create temporary table temp10 \n" +
                "select a1.* from temp09 a1, temp099 a2\n" +
                "where a1.registno = a2.registno\n" +
                "and a1.endcasedate <= a2.endcasedate\n" +
                "group by a1.registno, a1.endcasedate\n" +
                "having count(*) <= 1;");
        picczq.update("create index registno on temp10(registno);");
        picczq.update("update temp01 set endcasedate = ifnull((select a.endcasedate from temp10 a\n" +
                "where a.registno = temp01.registno),'1900-01-01 00:00:00');");
        picczq.update("update temp01 set endcasedate = null\n" +
                "where endcasedate = '1900-01-01';");
        picczq.update("update temp01 set flag3 = '已结案'\n" +
                "where exists (select * from  cl_prplclaim lclaim\n" +
                "where lclaim.registno = temp01.registno\n" +
                "and lclaim.endcasedate is not null);");
        picczq.update("create index  policyno on temp01(policyno);");
        picczq.update("create temporary table temp11 \n" +
                "select  a.*,b.agentcode,left(b.comcode,10) as comcode,b.handler1code,b.insuredcode,b.insuredname,left(a.brandname,4) as brandname1,\n" +
                "b.startdate,b.enddate, e.damagedate, e.reportdate, e.reportornumber,f.monopolycode, f.monopolyname\n" +
                "from temp01 a left join ( cl_prplcmain b  left join cx_prpcmain g\n" +
                "on  b.policyno = g.policyno left join cx_prpcitem_car f\n" +
                "on g.proposalno = f.proposalno)\n" +
                "on a.policyno = b.policyno\n" +
                "and a.registno = b.registno \n" +
                "left join cl_prplregist e\n" +
                "on a.registno = e.registno;");
        picczq.update("create index   registno on  temp11(registno); ");
        picczq.update("create temporary table temp12 \n" +
                "select a.*, b.checknature, c.checker1, c.checker2 \n" +
                "from temp11 a  left join cl_prplcheck b\n" +
                "on a.registno = b.registno\n" +
                "and b.validflag = '1'\n" +
                "left join  cl_prplchecktask c\n" +
                "on a.registno = c.registno\n" +
                "and c.validflag = '1';");
        picczq.update("create temporary table temp13 \t\n" +
                "select a.*,cast('' as nchar(20)) as licenseno0,cast('' as nchar(100)) as brandname0,cast('' as nchar(20)) as useyears0\n" +
                "from temp12 a;");
        picczq.update("create index registno on temp13(registno);");
        picczq.update("update temp13 set licenseno0 = (select max(licenseno) from cl_prplcitemcar\n" +
                "where temp13.registno = cl_prplcitemcar.registno);");
        picczq.update("update temp13 set brandname0 = (select max(brandname) from cl_prplcitemcar\n" +
                "where temp13.registno = cl_prplcitemcar.registno);");
        picczq.update("update temp13 set useyears0 = (select max(useyears) from cl_prplcitemcar\n" +
                "where temp13.registno = cl_prplcitemcar.registno);");

        //肇庆

        List<Map<String, Object>> result1 = picczq.queryForList("select * from temp13 order by enddeflossdate desc;");
        //车商

        List<Map<String, Object>> result2 = picczq.queryForList("select * from temp13\n" +
                "where comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342') \n" +
                "order by UNDERWRITEENDDATE desc;");
        //端州

        List<Map<String, Object>> result3 = picczq.queryForList("select * from temp13\n" +
                "where comcode in ('44129311','44120100','44120313','44129301','44129302','44129303','44129315','44129316','44129317','44129318','44129319','44129321','44129324','44129329','44129330','44129338')\n" +
                "order by enddeflossdate desc;");
        //鼎湖
        List<Map<String, Object>> result4 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like '441203%' and comcode <>'44120313'\n" +
                "order by enddeflossdate desc;");
        //四会
        List<Map<String, Object>> result5 = picczq.queryForList("select * from temp13\n" +
                "where comcode  like  '441284%'" +
                "or comcode in ('44129323','44129325')" +
                "order by enddeflossdate desc;");
        //广宁
        List<Map<String, Object>> result6 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441223%'\n" +
                "order by enddeflossdate desc;");
        //怀集

        List<Map<String, Object>> result7 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441224%'\n" +
                "order by enddeflossdate desc;");
        //高要

        List<Map<String, Object>> result8 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441283%'\n" +
                "order by enddeflossdate desc;");
        //德庆
        List<Map<String, Object>> result9 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441226%'\n" +
                "order by enddeflossdate desc;");
        //封开
        List<Map<String, Object>> result10 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441225%'\n" +
                "order by enddeflossdate desc;");

        picczq.update("drop table temp011_tmp;");
        picczq.update("drop table temp011;");
        picczq.update("drop table temp01;");
        picczq.update("drop table temp02;");
        picczq.update("drop table temp04;");
        picczq.update("drop table temp055;");
        picczq.update("drop table temp05;");
        picczq.update("drop table temp06;");
        picczq.update("drop table temp077;");
        picczq.update("drop table temp07;");
        picczq.update("drop table temp08;");
        picczq.update("drop table temp09;");
        picczq.update("drop table temp099;");
        picczq.update("drop table temp10;");
        picczq.update("drop table temp11;");
        picczq.update("drop table temp12;");
        picczq.update("drop table temp13;");

        //肇庆
        exportToExcel.hesun1(result1,"肇庆");

        //车商
        exportToExcel.hesun1(result2,"车商");

        //端州
        exportToExcel.hesun1(result3,"端州");

        //鼎湖
        exportToExcel.hesun1(result4,"鼎湖");

        //四会
        exportToExcel.hesun1(result5,"四会");

        //广宁
        exportToExcel.hesun1(result6,"广宁");

        //怀集
        exportToExcel.hesun1(result7,"怀集");

        //高要
        exportToExcel.hesun1(result8,"高要");

        //德庆
        exportToExcel.hesun1(result9,"德庆");

        //封开
        exportToExcel.hesun1(result10,"封开");




        return new ResultMap(1,"",true,"","");
    }
    //结案
    public ResultMap jiean(){
        picczq.update("create temporary table jiean\n" +
                "select a2.comcode,a0.registno,b0.claimno,a2.lflag,b1.licenseno,b1.brandname,b0.policyno,b0.insuredname,b0.agentcode,b0.startdate,\n" +
                "b0.enddate,b0.businessnature,a0.damagedate,a0.damageaddress,a0.reportdate,c.endcasedate,a0.reportornumber,a2.repairfactorycode,a2.repairfactoryname,\n" +
                "a2.handlername,a2.deflossdate,a2.underwritename,truncate(a2.sumverilossfee,2),truncate(a2.sumlossfee,2),\n" +
                "cast('' as nchar(100)) as checker1,b.handlercode,\n" +
                "cast('' as nchar(100)) as handname,b.handler1code,\n" +
                "cast('' as nchar(100)) as hand1name,\n" +
                "cast('' as nchar(100)) as agentname,a2.underwriteenddate\n" +
                "from  cl_prplregist a0,cl_prpldeflossmain a2,cl_prplcmain b0,cl_prplcitemcar b1,cx_prpcmain b,cl_prplclaim c\n" +
                "where substring(a2.underwriteflag,1,1) in ('1','3') \n" +
                "and a2.validflag='1' \n" +
                "and c.endcasedate >=?\n" +
                "and c.endcasedate < ?\n" +
                "and c.casetype='2'\n" +
                "and c.endcasedate is not null\n" +
                "and a2.registno=b0.registno \n" +
                "and a2.riskcode=b0.riskcode \n" +
                "and b0.id=b1.prplcmainid \n" +
                "and a0.registno=c.registno\n" +
                "and a2.registno=a0.registno\n" +
                "and b0.policyno=b.policyno;",todayminone,today);
        picczq.update("update jiean set handname = (select max(username) from utiihandler \n" +
                "where jiean.handlercode = utiihandler.usercode);");
        picczq.update("update jiean set checker1 = (select max(checker1) from cl_prplchecktask \n" +
                "where jiean.registno = cl_prplchecktask.registno);");
        picczq.update("update jiean set hand1name = (select max(username) from utiihandler \n" +
                "where jiean.handler1code = usercode);");
        picczq.update("update jiean set agentname = (select agentname from prpdagent \n" +
                "where jiean.agentcode = prpdagent.agentcode);");
        //肇庆

        List<Map<String, Object>> result1 = picczq.queryForList("select * from jiean order by endcasedate desc;");
        //车商

//
        List<Map<String, Object>> result2 = picczq.queryForList("select * from jiean\n" +
                "where comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342') \n" +
                "order by endcasedate desc;");
        //端州

//
        List<Map<String, Object>> result3 = picczq.queryForList("select * from jiean\n" +
                "where comcode in ('44129311','44120100','44120313','44129301','44129302','44129303','44129315','44129316','44129317','44129318','44129319','44129321','44129324','44129329','44129330','44129338')\n" +
                "order by endcasedate desc;");
        //鼎湖

//
        List<Map<String, Object>> result4 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like '441203%' and comcode <>'44120313'\n" +
                "order by endcasedate desc;");
        //四会


//
        List<Map<String, Object>> result5 = picczq.queryForList("select * from jiean\n" +
                "where comcode  like  '441284%'" +
                "or comcode in ('44129323','44129325')" +
                "order by endcasedate desc;");
        //广宁

        List<Map<String, Object>> result6 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441223%'\n" +
                "order by endcasedate desc;");
        //怀集

        List<Map<String, Object>> result7 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441224%'\n" +
                "order by endcasedate desc;");
        //高要

        List<Map<String, Object>> result8 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441283%'\n" +
                "order by endcasedate desc;");
        //德庆

        List<Map<String, Object>> result9 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441226%'\n" +
                "order by endcasedate desc;");
        //封开

        List<Map<String, Object>> result10 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441225%'\n" +
                "order by endcasedate desc;");
        picczq.update("drop table jiean");
        //肇庆
        exportToExcel.jiean(result1,"肇庆");

        //车商
        exportToExcel.jiean(result2,"车商");

        //端州
        exportToExcel.jiean(result3,"端州");

        //鼎湖
        exportToExcel.jiean(result4,"鼎湖");

        //四会
        exportToExcel.jiean(result5,"四会");

        //广宁
        exportToExcel.jiean(result6,"广宁");

        //怀集
        exportToExcel.jiean(result7,"怀集");

        //高要
        exportToExcel.jiean(result8,"高要");

        //德庆
        exportToExcel.jiean(result9,"德庆");

        //封开
        exportToExcel.jiean(result10,"封开");

        return null;
    }
    //保呗商城
    public ResultMap bbsc(){
        picczq.update("create temporary table temp1(\n" +
                "select cast(null as char(120))comname,a.comcode,cast(null as char(120))comcodename,a.proposalno,a.policyno,a.sumpremium,\n" +
                "a.sumtaxpremium,a.sumnetpremium,a.agentcode,cast(null as char(120)) agentname,b.clausetype,\n" +
                "b.carkindcode,cast(null as char(120))carkindname,a.startdate,a.enddate,b.licenseno,a.businessnature,cast(null as char(120))businessnaturename,\n" +
                "a.operatedate,a.riskcode,b.monopolycode,b.monopolyname,a.handler1code,cast(null as char(120))handler1name,\n" +
                "a.handlercode,cast(null as char(120))handlername,cast(null as datetime) policynotime,a.projectcode,\n" +
                "c.clausecode,c.clausename,c.clauses,b.newcarflag,cast(null as char(22))lastpolicyno,\n" +
                "b.usenaturecode,cast(null as char(20)) usenaturename,cast(null as char(20))xinxuzhuan,\n" +
                "cast(null as char(120)) insuredname,cast(null as char(120)) bbrinsuredname,cast(null as char(4))insuredtype,cast(0.0 as decimal(65,4))jifen\n" +
                "from cx_prpcmain a,cx_prpcitem_car b,cx_prpcengage c\n" +
                "where c.clausecode in ('997468','997469','997470','997471','997472','9908983','9908985','9908987','9908988','9908989','9908991','9908992','000121')\n" +
                "and a.operatedate >= ?\n" +
                "and a.policyno is not null\n" +
                "and a.policyno != ''\n" +
                "and a.proposalno = b.proposalno\n" +
                "and a.proposalno = c.proposalno\n" +
                "); ",todayminone);
        picczq.update("create index proposalno on temp1(proposalno);");
        picczq.update("UPDATE temp1\n" +
                "    SET jifen = CASE temp1.clausecode \n" +
                "        when '997468' then '200'\n" +
                "        when '997469' then '400'\n" +
                "        when '997470' then '600'\n" +
                "        when '997471' then '800'\n" +
                "        when '997472' then '1000'\n" +
                "        when '9908983' then floor(sumnetpremium * 0.05)\n" +
                "        when '9908985' then floor(sumnetpremium * 0.10)\n" +
                "        when '9908987' then floor(sumnetpremium * 0.15)\n" +
                "        when '9908988' then floor(sumnetpremium * 0.20)\n" +
                "        when '9908989' then floor(sumnetpremium * 0.25)\n" +
                "        when '9908991' then floor(sumnetpremium * 0.30)\n" +
                "        when '9908992' then floor(sumnetpremium * 0.35)\n" +
                "    END\n" +
                "WHERE temp1.clausecode IN ('997468','997469','997470','997471','997472','9908983','9908985','9908987','9908988','9908989','9908991','9908992');");

        picczq.update("update temp1 set comname = '车商'\n" +
                "        where  comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342');");
        picczq.update("update temp1 set comname = '端州'\n" +
                "        where comcode in ('44120100','44129301','44129302','44129303','44129311','44129315','44129317','44129318','44129319',\n" +
                "              '44129321','44129329','44129330','44129338','44129343','44129324','44120313');");
        picczq.update("update temp1 set comname = '鼎湖'\n" +
                "        where comcode like '441203%' and comcode <>'44120313';");
        picczq.update("update temp1 set comname = '广宁'\n" +
                "        where comcode like '441223%';");
        picczq.update("update temp1 set comname = '怀集'\n" +
                "        where comcode like '441224%';");
        picczq.update("update temp1 set comname = '封开'\n" +
                "        where comcode like '441225%';");
        picczq.update("update temp1 set comname = '德庆'\n" +
                "        where comcode like '441226%';");
        picczq.update("update temp1 set comname = '高要'\n" +
                "        where comcode like '441283%';");
        picczq.update("update temp1 set comname = '四会'\n" +
                "        where comcode like '441284%';");
        picczq.update("update temp1 set comname = '大旺'\n" +
                "        where comcode in ('44129323','44129325');");
        picczq.update("update temp1 set comname = '电销'\n" +
                "        where comcode in ('44129400');");
        picczq.update("create temporary table temp2(\n" +
                "select proposalno,policyno,oldpolicyno\n" +
                "from cx_prpcrenewal \n" +
                "where proposalno in (select proposalno from temp1));");
        picczq.update("update temp1 set xinxuzhuan = '转保';");
        picczq.update("update temp1 set temp1.xinxuzhuan = '新保' \n" +
                "where newcarflag = 1;");
        picczq.update("update temp1,temp2 set temp1.xinxuzhuan = '续保' \n" +
                "where temp1.proposalno = temp2.proposalno;");
        picczq.update("update temp1,temp2 set  lastpolicyno = oldpolicyno\n" +
                "where temp1.proposalno  = temp2.proposalno;");
        picczq.update("update temp1,prpdcompany set comcodename = comcname\n" +
                "where temp1.comcode = prpdcompany.comcode;");
        picczq.update("update temp1,prpdagent set temp1.agentname = prpdagent.agentname\n" +
                "where temp1.agentcode = prpdagent.agentcode;");
        picczq.update("update temp1\n" +
                "    set carkindname = case temp1.carkindcode \n" +
                "        when 'a01' then '客车'\n" +
                "        when 'b01' then '货车'\n" +
                "        when 'b02' then '半挂牵引车'\n" +
                "        when 'b11' then '三轮汽车'\n" +
                "        when 'b12' then '低速货车'\n" +
                "        when 'b13' then '客货两用车'\n" +
                "        when 'b21' then '自卸货车'\n" +
                "        when 'b91' then '货车挂车'\n" +
                "        when 'c01' then '油罐车'\n" +
                "        when 'c02' then '气罐车'\n" +
                "        when 'c03' then '液罐车'\n" +
                "        when 'c04' then '冷藏车'\n" +
                "        when 'c11' then '罐车挂车'\n" +
                "        when 'c20' then '推土车'\n" +
                "        when 'c22' then '清障车'\n" +
                "        when 'c23' then '清扫车'\n" +
                "        when 'c24' then '清洁车'\n" +
                "        when 'c25' then '起重车'\n" +
                "        when 'c26' then '装卸车'\n" +
                "        when 'c27' then '升降车'\n" +
                "        when 'c28' then '混凝土搅拌车'\n" +
                "        when 'c29' then '挖掘车'\n" +
                "        when 'c30' then '专业拖车'\n" +
                "        when 'c31' then '特种车二挂车'\n" +
                "        when 'c39' then '特种车二类其它'\n" +
                "        when 'c41' then '电视转播车'\n" +
                "        when 'c42' then '消防车'\n" +
                "        when 'c43' then '医疗车'\n" +
                "        when 'c44' then '油汽田操作用车'\n" +
                "        when 'c45' then '压路车'\n" +
                "        when 'c46' then '矿山车'\n" +
                "        when 'c47' then '运钞车'\n" +
                "        when 'c48' then '救护车'\n" +
                "        when 'c49' then '监测车'\n" +
                "        when 'c50' then '雷达车'\n" +
                "        when 'c51' then 'x光检查车'\n" +
                "        when 'c52' then '电信抢修车/电信工程车'\n" +
                "        when 'c53' then '电力抢修车/电力工程车'\n" +
                "        when 'c54' then '专业净水车'\n" +
                "        when 'c55' then '保温车'\n" +
                "        when 'c56' then '邮电车'\n" +
                "        when 'c57' then '警用特种车'\n" +
                "        when 'c58' then '混凝土泵车'\n" +
                "        when 'c61' then '特种车三类挂车'\n" +
                "        when 'c69' then '特种车三类其它'\n" +
                "        when 'c90' then '集装箱拖头'\n" +
                "        when 'd01' then '摩托车'\n" +
                "        when 'd02' then '正三轮摩托车'\n" +
                "        when 'd03' then '侧三轮摩托车'\n" +
                "        when 'e01' then '拖拉机'\n" +
                "        when 'e11' then '联合收割机'\n" +
                "        when 'e12' then '变形拖拉机/其它'\n" +
                "        when 'z99' then '其它车辆'\n" +
                "    end\n" +
                "where temp1.carkindcode in ('a01','b01','b02','b11','b12','b13','b21','b91','c01','c02','c03','c04','c11','c20',\n" +
                "'c22','c23','c24','c25','c26','c27','c28','c29','c30','c31','c39','c41','c42','c43','c44','c45','c46','c47','c48',\n" +
                "'c49','c50','c51','c52','c53','c54','c55','c56','c57','c58','c61','c69','c90','d01','d02','d03','e01','e11','e12','z99'\n" +
                ");");
        picczq.update("update temp1\n" +
                "    set businessnaturename = case temp1.businessnature\n" +
                "        when '1' then '个人代理业务'\n" +
                "        when '0' then '传统直销业务'\n" +
                "        when '2' then '专业代理业务'\n" +
                "        when '3' then '兼业代理业务'\n" +
                "        when '4' then '经纪业务'\n" +
                "        when '53' then '网上业务'\n" +
                "    end\n" +
                "where temp1.businessnature in ('1','0', '2','3','4','53');");
        picczq.update("update temp1,utiihandler set handler1name = username\n" +
                "where utiihandler.usercode = temp1.handler1code;");
        picczq.update("update temp1,utiihandler set handlername = username\n" +
                "where utiihandler.usercode = temp1.handlercode;");
        picczq.update("update temp1,prptime set temp1.policynotime = prptime.operatetimeforhis\n" +
                "where temp1.policyno = prptime.certino\n" +
                "and prptime.updatetype = 'c01';");
        picczq.update("update temp1\n" +
                "    set usenaturename = case temp1.usenaturecode \n" +
                "        when '000' then '不区分营业非营业'\n" +
                "        when '111' then '出租、租赁'\n" +
                "        when '112' then '城市公交'\n" +
                "        when '113' then '公路客运'\n" +
                "        when '114' then '旅游客运'\n" +
                "        when '120' then '营业货车'\n" +
                "        when '121' then '营业挂车'\n" +
                "        when '180' then '运输型拖拉机'\n" +
                "        when '190' then '其它营业车辆'\n" +
                "        when '211' then '家庭自用汽车'\n" +
                "        when '212' then '非营业企业客车'\n" +
                "        when '213' then '非营业机关、事业团体客车'\n" +
                "        when '220' then '非营业货车'\n" +
                "        when '221' then '非营业挂车'\n" +
                "        when '280' then '兼用型拖拉机'\n" +
                "        when '290' then '其它非营业车辆'\n" +
                "    end\n" +
                "where temp1.usenaturecode in ('000','111','112','113','114','120','121','180','190','211','212','213','220','221','280','290');");
        picczq.update("update temp1,cx_prpcinsured set temp1.insuredtype = '个人'\n" +
                "where cx_prpcinsured.proposalno = temp1.proposalno\n" +
                "and cx_prpcinsured.insuredtype = 1;");
        picczq.update("update temp1,cx_prpcinsured set temp1.insuredtype = '团体'\n" +
                "where cx_prpcinsured.proposalno = temp1.proposalno\n" +
                "and cx_prpcinsured.insuredtype = 2;");
        picczq.update("update temp1 a,cx_prpcinsured b set a.insuredname = b.insuredname\n" +
                "where a.proposalno = b.proposalno\n" +
                "and substring(b.insuredflag,1,1) = 1;");
        picczq.update("update temp1 a,cx_prpcinsured b set a.bbrinsuredname = b.insuredname\n" +
                "where a.proposalno = b.proposalno\n" +
                "and substring(b.insuredflag,2,1) = 1;");
        picczq.update("create temporary table temp3(\n" +
                "select a.proposalno,a.serialno,cast(null as char(120))mainsellfeerate,sellfeerate,\n" +
                "cast(null as char(120))summainsellfeerate,cast(null as char(120))sumsellfeerate,\n" +
                "cast(null as char(120))maxcostrate,a.mainflag \n" +
                "from cx_prpcseller a,temp1 b\n" +
                "where b.proposalno = a.proposalno\n" +
                ");");
        picczq.update("create temporary table temp4(\n" +
                "select sum(sellfeerate)sellfeerate,proposalno,mainflag\n" +
                "from temp3\n" +
                "group by proposalno,mainflag\n" +
                ");\n");
        picczq.update("update temp3 a,temp4 b\n" +
                "set mainsellfeerate = b.sellfeerate,summainsellfeerate = b.sellfeerate\n" +
                "where  b.mainflag = 1\n" +
                "and a.proposalno = b.proposalno;");
        picczq.update("update temp3 a,temp4 b\n" +
                "set sumsellfeerate = b.sellfeerate\n" +
                "where  b.mainflag = 0\n" +
                "and a.proposalno = b.proposalno;");
        picczq.update("create temporary table temp5(\n" +
                "select cast(null as char(120))proposalno,a.policyno,max(a.maxcostrate)maxcostrate from scmsdocfeedetail a,temp1 b\n" +
                "where a.certino = b.policyno\n" +
                "and a.certitype = 'p'\n" +
                "group by a.policyno\n" +
                ");");
        picczq.update("update temp5 a,temp1 b set a.proposalno = b.proposalno\n" +
                "where a.policyno = b.policyno;");
        picczq.update("update temp3 a,temp5 b set a.maxcostrate = b.maxcostrate\n" +
                "where a.proposalno = b.proposalno;");

        List<Map<String, Object>> result  = picczq.queryForList("select a.comname,a.comcodename,a.policyno,a.sumpremium,\n" +
                "a.sumtaxpremium,a.sumnetpremium,a.agentname,a.clausetype,\n" +
                "a.carkindname,a.startdate,a.enddate,a.licenseno,a.businessnaturename,\n" +
                "a.operatedate,a.riskcode,a.monopolycode,a.monopolyname,a.handler1name,\n" +
                "a.handlername,a.policynotime,a.projectcode,a.insuredname,a.bbrinsuredname,\n" +
                "a.xinxuzhuan,b.mainsellfeerate,b.sellfeerate,b.summainsellfeerate,\n" +
                "b.sumsellfeerate,b.maxcostrate,a.lastpolicyno,a.usenaturename,\n" +
                "a.insuredtype,a.clausecode,a.clausename,a.clauses\n" +
                "from temp1 a,temp3 b\n" +
                "where a.proposalno = b.proposalno;");

        List<Map<String, Object>> result2 = picczq.queryForList("select sum(sumnetpremium)sumnetpremium,clausecode,clausename,sum(jifen) sumjifen,count(*)count from temp1\n" +
                "group by clausecode");
        List<Map<String, Object>> total = picczq.queryForList("select sum(jifen)total from temp1\n");
        picczq.update("drop table temp1");
        picczq.update("drop table temp2");
        picczq.update("drop table temp3");
        picczq.update("drop table temp4");
        picczq.update("drop table temp5");


        exportToExcel.bbsc(result);
        exportToExcel.bbsc2(result2,total);

        return  new ResultMap(1,"",true,"","");

    }
    //报案月
    public ResultMap monthbaoan(){
        picczq.update("create temporary table baoan\n" +
                "select  a.makecom,c.licenseno,c.policyno,a.registno,a.reportdate,a.reportorname,a.reportornumber,a.damagedate,a.damagehour,a.damageaddress,a.damagename,\n" +
                "b.lflag,b.checker1,truncate(b.sumestimatefee,2),b.firstsiteflag \n" +
                "from cl_prplregist a,\n" +
                "cl_prplregistsummary c,\n" +
                "cl_prplchecktask b\n" +
                "where a.makecom like '4412%' \n" +
                "and a.validflag='1'" +
                "and a.reportdate >= ? \n" +
                "and a.reportdate < ? \n" +
                "and a.registno=c.registno \n" +
                "and a.registno=b.registno;",lastmonth,month);
        picczq.update("create index policyno on baoan(policyno);");
        picczq.update("create temporary table baoan1\n" +
                "select a.*,b.handlercode,cast('' as nchar(100)) as handlername,b.handler1code,\n" +
                "cast('' as nchar(100)) as handler1name\n" +
                "from baoan a,\n" +
                "cx_prpcmain b\n" +
                "where a.policyno=b.policyno");
        picczq.update("update baoan1 set handlername = (select max(username) from utiihandler a \n" +
                "where baoan1.handlercode = a.usercode);");
        picczq.update("update baoan1 set handler1name = (select max(username) from utiihandler a \n" +
                "where baoan1.handler1code = a.usercode);");

        //肇庆
        List<Map<String, Object>> result1 = picczq.queryForList("select * from baoan1 order by reportdate desc;");
        //车商

        List<Map<String, Object>> result2 = picczq.queryForList("select * from baoan1\n" +
                "where makecom in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342') \n" +
                "order by reportdate desc;");
        //端州
        List<Map<String, Object>> result3 = picczq.queryForList("select * from baoan1\n" +
                "where makecom in ('44129311','44120100','44120313','44129301','44129302','44129303','44129315','44129316','44129317','44129318','44129319','44129321','44129324','44129329','44129330','44129338')\n" +
                "order by reportdate desc;");
        //鼎湖
        List<Map<String, Object>> result4 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like '441203%' and makecom <>'44120313'\n" +
                "order by reportdate desc;");
        //四会

        List<Map<String, Object>> result5 = picczq.queryForList("select * from baoan1\n" +
                "where makecom  like  '441284%'" +
                "or makecom in ('44129323','44129325')" +
                "order by reportdate desc;");
        //广宁

        List<Map<String, Object>> result6 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441223%'\n" +
                "order by reportdate desc;");
        //怀集

        List<Map<String, Object>> result7 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441224%'\n" +
                "order by reportdate desc;");
        //高要

        List<Map<String, Object>> result8 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441283%'\n" +
                "order by reportdate desc;");
        //德庆

        List<Map<String, Object>> result9 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441226%'\n" +
                "order by reportdate desc;");
        //封开

        List<Map<String, Object>> result10 = picczq.queryForList("select * from baoan1\n" +
                "where makecom   like  '441225%'\n" +
                "order by reportdate desc;");

        picczq.update("drop table baoan;");
        picczq.update("drop table baoan1;");

        exportToExcel.baoan2(result1,"肇庆");


        //车商
        exportToExcel.baoan2(result2,"车商");

        //端州
        exportToExcel.baoan2(result3,"端州");

        //鼎湖
        exportToExcel.baoan2(result4,"鼎湖");

        //四会
        exportToExcel.baoan2(result5,"四会");

        //广宁
        exportToExcel.baoan2(result6,"广宁");

        //怀集
        exportToExcel.baoan2(result7,"怀集");

        //高要
        exportToExcel.baoan2(result8,"高要");

        //德庆
        exportToExcel.baoan2(result9,"德庆");

        //封开
        exportToExcel.baoan2(result10,"封开");

        return null;
    }
    //送修月
    public ResultMap monthsongxiu(){
        picczq.update("create temporary table songxiu\n" +
                "select a2.comcode,a0.registno,b0.claimno,a2.lflag,b1.licenseno,b1.brandname,b0.policyno,b0.insuredname,b0.agentcode,b0.startdate,\n" +
                "b0.enddate,b0.businessnature,a0.damagedate,a0.damageaddress,a0.reportdate,a0.reportornumber,a2.repairfactorycode,a2.repairfactoryname,\n" +
                "a2.handlername,a2.deflossdate,a2.underwritename,truncate(a2.sumverilossfee,2),truncate(a2.sumlossfee,2),\n" +
                "cast('' as nchar(100)) as checker1,b.handlercode,\n" +
                "cast('' as nchar(100)) as handname,b.handler1code,\n" +
                "cast('' as nchar(100)) as hand1name,\n" +
                "cast('' as nchar(100)) as agentname,a2.underwriteenddate\n" +
                "from  cl_prplregist a0,cl_prpldeflossmain a2,cl_prplcmain b0,cl_prplcitemcar b1,cx_prpcmain b\n" +
                "where substring(a2.underwriteflag,1,1) in ('1','3') \n" +
                "and a2.validflag='1' \n" +
                "and a2.underwriteenddate >= ?\n" +
                "and a2.underwriteenddate < ?\n" +
                "and a2.registno=b0.registno \n" +
                "and a2.riskcode=b0.riskcode \n" +
                "and b0.id=b1.prplcmainid \n" +
                "and a2.registno=a0.registno\n" +
                "and b0.policyno=b.policyno;",lastmonth,month);
        picczq.update("update songxiu set handname = (select max(username) from utiihandler \n" +
                "where songxiu.handlercode = utiihandler.usercode);");
        picczq.update("update songxiu set checker1 = (select max(checker1) from cl_prplchecktask \n" +
                "where songxiu.registno = cl_prplchecktask.registno);");
        picczq.update("update songxiu set hand1name = (select max(username) from utiihandler \n" +
                "where songxiu.handler1code = usercode);");
        picczq.update("update songxiu set agentname = (select agentname from prpdagent \n" +
                "where songxiu.agentcode = prpdagent.agentcode);");

        //肇庆

        List<Map<String, Object>> result1 = picczq.queryForList("select * from songxiu order by UNDERWRITEENDDATE desc;");
        //车商


        List<Map<String, Object>> result2 = picczq.queryForList("select * from songxiu\n" +
                "where comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342') \n" +
                "order by UNDERWRITEENDDATE desc;");
        //端州
        List<Map<String, Object>> result3 = picczq.queryForList("select * from songxiu\n" +
                "where comcode in ('44129311','44120100','44120313','44129301','44129302','44129303','44129315','44129316','44129317','44129318','44129319','44129321','44129324','44129329','44129330','44129338')\n" +
                "order by UNDERWRITEENDDATE desc;");
        //鼎湖

        List<Map<String, Object>> result4 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like '441203%' and comcode <>'44120313'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //四会


        List<Map<String, Object>> result5 = picczq.queryForList("select * from songxiu\n" +
                "where comcode  like  '441284%'" +
                "or comcode in ('44129323','44129325')" +
                "order by UNDERWRITEENDDATE desc;");
        //广宁

        List<Map<String, Object>> result6 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441223%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //怀集
        List<Map<String, Object>> result7 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441224%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //高要

        List<Map<String, Object>> result8 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441283%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //德庆


        List<Map<String, Object>> result9 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441226%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        //封开

        List<Map<String, Object>> result10 = picczq.queryForList("select * from songxiu\n" +
                "where comcode   like  '441225%'\n" +
                "order by UNDERWRITEENDDATE desc;");
        picczq.update("drop table songxiu");
        //肇庆
        exportToExcel.songxiu2(result1,"肇庆");
        //车商
        exportToExcel.songxiu2(result2,"车商");
        //端州
        exportToExcel.songxiu2(result3,"端州");
        //鼎湖
        exportToExcel.songxiu2(result4,"鼎湖");
        //四会
        exportToExcel.songxiu2(result5,"四会");
        //广宁
        exportToExcel.songxiu2(result6,"广宁");
        //怀集
        exportToExcel.songxiu2(result7,"怀集");
        //高要
        exportToExcel.songxiu2(result8,"高要");;
        //德庆
        exportToExcel.songxiu2(result9,"德庆");
        //封开
        exportToExcel.songxiu2(result10,"封开");

        return null;
    }
    //核损月
    public ResultMap monthhesun(){
        picczq.update("create temporary table temp011_tmp \n" +
                "select a.sumlossfee,a.sumverilossfee,a.cetainlosstype,a.repairfactorytype,a.registno,a.policyno,a.inputtime,a.enddeflossdate, a.finalhandlername,a.underwritecode,\n" +
                "a.underwritename,a.underwriteenddate,a.sumverichgcompfee,a.sumverirepairfee,\n" +
                "(case when left(a.underwriteflag,1) in ('1','3') \n" +
                "then '已核损' else '未核损' end) as  flag1,a.examfactorycode,a.examfactoryname,a.repairfactorycode,a.repairfactoryname,a.prplthirdpartid,0 as reclaim, d.handlercode as reclaimcode, \n" +
                "d.handlername as reclaimname, d.inputtime as reclaiminputtime,d.inputtime as reclaimoutputtime,d.id as reclaimid, a.id as lossmainid, '本地保单' as lflag\n" +
                "from  cl_prpldeflossmain a \n" +
                "inner join cl_prplregistsummary b\n" +
                "on  a.registno = b.registno\n" +
                "and a.riskcode = b.riskcode\n" +
                "and a.enddeflossdate >=?\n" +
                "and a.enddeflossdate < ?\n" +
                "and b.comcode like '4412%' \n" +
                "left join  cl_prplreclaim d \n" +
                "on  a.registno = d.registno\n" +
                "and a.id = d.lossmainid\n" +
                "and d.validflag = '1';",lastmonth,month);
        picczq.update("insert into temp011_tmp\n" +
                "select  a.sumlossfee,a.sumverilossfee,a.cetainlosstype,a.repairfactorytype,a.registno,a.policyno,a.inputtime,a.enddeflossdate,a.finalhandlername,a.underwritecode,\n" +
                "a.underwritename,a.underwriteenddate,a.sumverichgcompfee,a.sumverirepairfee,(case when left(a.underwriteflag,1)  in ('1','3') \n" +
                "then '已核损' else '未核损' end) as  flag1,a.examfactorycode,a.examfactoryname,a.repairfactorycode,a.repairfactoryname,a.prplthirdpartid,0 as reclaim, d.handlercode as reclaimcode, \n" +
                "d.handlername as reclaimname, d.inputtime as reclaiminputtime, d.inputtime as reclaimoutputtime, \n" +
                "d.id as reclaimid, a.id as ossmainid, '外地保单' as lflag\n" +
                "from  cl_prpldeflossmain a \n" +
                "inner join cl_prplregistsummary b\n" +
                "on a.registno = b.registno\n" +
                "and a.riskcode = b.riskcode\n" +
                "and a.enddeflossdate >=?\n" +
                "and a.enddeflossdate < ?\n" +
                "and (b.comcode not like '4412%' and a.makecom like  '4412%')                \n" +
                "left join  cl_prplreclaim d \n" +
                "on  a.registno = d.registno\n" +
                "and a.id = d.lossmainid\n" +
                "and d.validflag = '1';",lastmonth,month);
        picczq.update("create temporary table temp011 \t\n" +
                "select  a.sumlossfee,a.sumverilossfee,a.cetainlosstype,a.repairfactorytype,\n" +
                "cast('' as char(20)) as repairfactorytypename,\n" +
                "a.registno,a.lflag,a.policyno,\n" +
                "a.inputtime,a.enddeflossdate, a.finalhandlername,a.underwritecode,a.underwritename,a.underwriteenddate,a.sumverichgcompfee,a.sumverirepairfee,a.flag1,a.examfactorycode,a.examfactoryname,\n" +
                "a.repairfactorycode,a.repairfactoryname,a.prplthirdpartid,reclaim, reclaimcode,reclaimname, reclaiminputtime, reclaimoutputtime, reclaimid, lossmainid\n" +
                "from temp011_tmp a;");
        picczq.update("create index inx_3a01a_01 on temp011(prplthirdpartid);");
        picczq.update("create index inx_3a01a_011 on temp011(lossmainid);");
        picczq.update("create index registno on temp011(registno);");
        picczq.update("create index reclaimid on temp011(reclaimid);");
        picczq.update("update temp011\n" +
                "set reclaiminputtime = ((select max(indate)from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.reclaimid\n" +
                "and a.nodeid = '25')),\n" +
                "reclaimoutputtime = ((select max(outdate)from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.reclaimid\n" +
                "and a.nodeid = '25'));");
        picczq.update("update temp011\n" +
                "set reclaimcode= ((select max(a.usercode) from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.lossmainid\n" +
                "and a.nodeid = '25'\n" +
                "and a.valid = '1')),\n" +
                "reclaiminputtime= ((select max(indate)\n" +
                "from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.lossmainid\n" +
                "and a.nodeid = '25'\n" +
                "and a.valid = '1')),\n" +
                "reclaimoutputtime = ((select  max(outdate)\n" +
                "from  cl_prplbpmmain a\n" +
                "where a.mainno = temp011.registno\n" +
                "and a.businessid = temp011.lossmainid\n" +
                "and a.nodeid = '25'\n" +
                "and a.valid = '1'))\n" +
                "where reclaimcode is null;");
        picczq.update("update temp011\n" +
                "set reclaimname = (select username from prpduser a\n" +
                "where  a.usercode = temp011.reclaimcode);");
        picczq.update("update temp011\n" +
                "set reclaim = 1\n" +
                "where reclaimcode is not null;");
        picczq.update("create index inx_3a01a_02 on temp011(registno)");
        picczq.update("create temporary table temp01\n" +
                "select a.*, b.id ,0 as count1,0 as count2,b.brandname,'                 ' as flag3,'非人伤' as  flag4,'        ' as flag5,\n" +
                "cast('' as nchar(20)) as opname1,\n" +
                "cast('' as nchar(20)) as opname2, \n" +
                "cast('' as nchar(20)) as opname3, \n" +
                "cast('' as nchar(20)) as opname4,\n" +
                "cast('1900-01-01 00:00:00' as datetime) as outdate1, \n" +
                "cast('1900-01-01 00:00:00' as datetime) as outdate2, \n" +
                "cast('1900-01-01' as date) as endcasedate\n" +
                "from temp011 a  left join cl_prpldeflossthirdparty b\n" +
                "on a.registno = b.registno \n" +
                "and a.prplthirdpartid = b.id;");
        picczq.update("create index inx_3a01a_03 on temp01(registno);");
        picczq.update("delete from temp01\n" +
                "where exists (select * from  cl_prplregist reg\n" +
                "where reg.registno = temp01.registno\n" +
                "and reg.cancelflag = '1');");
        picczq.update("create temporary table temp02 \n" +
                "select * from  prpduser\n" +
                "where comcode like '4412%';");
        picczq.update("create index inx_tmp02_01 on temp02(usercode)");
        picczq.update("update temp01 \n" +
                "set opname1 = ifnull((select max(b.username) from  cl_prplcompensate a,temp02 b \n" +
                "where a.registno = temp01.registno\n" +
                "and a.operatorcode = b.usercode\n" +
                "and (a.underwriteflag is null or left(a.underwriteflag,1)  not in('7','8'))), '        ')\n" +
                "where flag1 = '已核损';");
        picczq.update("update temp01\n" +
                "set opname2 = (select max(underwritename) from  cl_prplcompensate a\n" +
                "where a.registno = temp01.registno\n" +
                "and (left(a.underwriteflag,1)  in ('1','3')))\n" +
                "where opname1 is not null;");
        picczq.update("update temp01\n" +
                "set flag3 = '已理算未核赔'\n" +
                "where 0 < (select count(*) from  cl_prplcompensate a\n" +
                "where a.registno = temp01.registno\n" +
                "and (a.underwriteflag is null or left(a.underwriteflag,1) in ('0','2','9')));");
        picczq.update("update temp01\n" +
                "set flag3 = '已理算已核赔'\n" +
                "where 0 < (select count(*) from  cl_prplcompensate a\n" +
                "where a.registno = temp01.registno\n" +
                "and (left(a.underwriteflag,1) in ('1','3')))\n" +
                "and flag3 <> '已理算未核赔';");
        picczq.update("update temp01\n" +
                "set flag3 = '已核损未理算'\n" +
                "where flag1= '已核损' and flag3 not like '已理算%';");
        picczq.update("update temp01\n" +
                "set flag3 = '已定损未核损'\n" +
                "where flag1= '未核损';");
        picczq.update("update temp01\n" +
                "set count1 = (select count(*) from cl_prplcomponent b\n" +
                "where b.prpldeflossmainid = temp01.id\n" +
                "and b.registno = temp01.registno\n" +
                "and b.validflag = '1')\n" +
                "where 1 = 1;");
        picczq.update("update temp01\n" +
                "set count2 = (select count(*) from cl_prplrepairfee b\n" +
                "where b.prpldeflossmainid = temp01.id\n" +
                "and b.registno = temp01.registno\n" +
                "and b.validflag = '1');");
        picczq.update("create index  registno on temp01(registno);");
        picczq.update("update temp01\n" +
                "set flag4 = '人伤案'\n" +
                "where 0 < (select count(*) from  cl_prplbpmmain a\n" +
                "where a.mainno = temp01.registno\n" +
                "and a.nodeid = 10\n" +
                "and a.cancelstate='0'\n" +
                "and a.valid ='1');");
        picczq.update("update temp01\n" +
                "set flag5 = ifnull((select max(a.damagecasename)\n" +
                "from  cl_prplcheck a\n" +
                "where a.registno = temp01.registno),'        ');");
        picczq.update("create temporary table temp04 \n" +
                "select a.registno, b.usercode, c.username \n" +
                "from temp01 a inner join  cl_prplbpmmain b\n" +
                "on  a.registno = b.mainno\n" +
                "and  b.nodeid = '14'\n" +
                "and b.valid = 1\n" +
                "left join  temp02 c\n" +
                "on b.usercode = c.usercode;");
        picczq.update("create index registno on temp04(registno);");
        picczq.update("update temp01 \n" +
                "set opname3 = ifnull((select max(a.username)\n" +
                "from temp04 a\n" +
                "where a.registno = temp01.registno),'');");
        picczq.update("create temporary table temp05 \n" +
                "select a.registno, b.usercode, b.outdate, c.username          \n" +
                "from temp01 a inner join cl_prplbpmmain b\n" +
                "on a.registno = b.mainno\n" +
                "and b.nodeid = 2\n" +
                "and b.prepnodeid = 4\n" +
                "and b.valid = 1\n" +
                "left join  temp02 c\n" +
                "on b.usercode = c.usercode\n" +
                "order by b.outdate;");
        picczq.update("create temporary table temp055 \n" +
                "select a.registno, b.usercode, b.outdate, c.username          \n" +
                "from temp01 a inner join cl_prplbpmmain b\n" +
                "on a.registno = b.mainno\n" +
                "and b.nodeid = 2\n" +
                "and b.prepnodeid = 4\n" +
                "and b.valid = 1\n" +
                "left join  temp02 c\n" +
                "on b.usercode = c.usercode\n" +
                "order by b.outdate;");
        picczq.update("create temporary table temp06 \n" +
                "select a1.* from temp05 a1, temp055 a2\n" +
                "where a1.registno = a2.registno\n" +
                "and a1.outdate >= a2.outdate\n" +
                "group by a1.registno, a1.usercode, a1.outdate, a1.username\n" +
                "having count(*) <= 1;");
        picczq.update("create index registno on temp06(registno);");
        picczq.update("update temp01 set opname4=ifnull\n" +
                "((select a.username from temp06 a\n" +
                "where a.registno = temp01.registno),'');");
        picczq.update("update temp01 set outdate1=ifnull((select  a.outdate  from temp06 a\n" +
                "where a.registno = temp01.registno),'1900-01-01 00:00:00');");
        picczq.update("update temp01 set outdate1 = null\n" +
                "where outdate1 = '1900-01-01 00:00:00';");
        picczq.update("create temporary table temp07 \n" +
                "select a.registno, b.outdate \n" +
                "from temp01 a,  cl_prplbpmmain b\n" +
                "where a.registno = b.mainno\n" +
                "and b.nodeid = 16\n" +
                "and b.valid = 1\n" +
                "and left(b.businesstype,1) = 'b'\n" +
                "order by b.outdate desc;");
        picczq.update("create temporary table temp077 \n" +
                "select a.registno, b.outdate \n" +
                "from temp01 a,  cl_prplbpmmain b\n" +
                "where a.registno = b.mainno\n" +
                "and b.nodeid = 16\n" +
                "and b.valid = 1\n" +
                "and left(b.businesstype,1) = 'b'\n" +
                "order by b.outdate desc;");
        picczq.update("create index registno on temp07(registno);");
        picczq.update("create index registno on temp077(registno);");
        picczq.update("create temporary table temp08 \n" +
                "select a1.* from temp07 a1, temp077 a2\n" +
                "where a1.registno = a2.registno\n" +
                "and a1.outdate <= a2.outdate\n" +
                "group by a1.registno, a1.outdate\n" +
                "having count(*) <= 1;");
        picczq.update("create index registno on temp08(registno);");
        picczq.update("update temp01 set outdate2 = ifnull((select a.outdate from temp08 a\n" +
                "where a.registno = temp01.registno),'1900-01-01 00:00:00');");
        picczq.update("update temp01 set outdate2 = null\n" +
                "where outdate2 = '1900-01-01 00:00:00';");
        picczq.update("update temp01\n" +
                "set policyno = (select max(policyno) from  cl_prplregistsummary a\n" +
                "where a.registno = temp01.registno)\n" +
                "where policyno is null;");
        picczq.update("create temporary table temp09 \n" +
                "select a.registno, b.endcasedate\n" +
                "from temp01 a,  cl_prplclaim b\n" +
                "where a.registno = b.registno\n" +
                "order by endcasedate desc;");
        picczq.update("create temporary table temp099 \n" +
                "select a.registno, b.endcasedate\n" +
                "from temp01 a,  cl_prplclaim b\n" +
                "where a.registno = b.registno\n" +
                "order by endcasedate desc;");
        picczq.update("create index registno on temp09(registno);");
        picczq.update("create index registno on temp099(registno);");
        picczq.update("create temporary table temp10 \n" +
                "select a1.* from temp09 a1, temp099 a2\n" +
                "where a1.registno = a2.registno\n" +
                "and a1.endcasedate <= a2.endcasedate\n" +
                "group by a1.registno, a1.endcasedate\n" +
                "having count(*) <= 1;");
        picczq.update("create index registno on temp10(registno);");
        picczq.update("update temp01 set endcasedate = ifnull((select a.endcasedate from temp10 a\n" +
                "where a.registno = temp01.registno),'1900-01-01 00:00:00');");
        picczq.update("update temp01 set endcasedate = null\n" +
                "where endcasedate = '1900-01-01';");
        picczq.update("update temp01 set flag3 = '已结案'\n" +
                "where exists (select * from  cl_prplclaim lclaim\n" +
                "where lclaim.registno = temp01.registno\n" +
                "and lclaim.endcasedate is not null);");
        picczq.update("create index  policyno on temp01(policyno);");
        picczq.update("create temporary table temp11 \n" +
                "select  a.*,b.agentcode,left(b.comcode,10) as comcode,b.handler1code,b.insuredcode,b.insuredname,left(a.brandname,4) as brandname1,\n" +
                "b.startdate,b.enddate, e.damagedate, e.reportdate, e.reportornumber,f.monopolycode, f.monopolyname\n" +
                "from temp01 a left join ( cl_prplcmain b  left join cx_prpcmain g\n" +
                "on  b.policyno = g.policyno left join cx_prpcitem_car f\n" +
                "on g.proposalno = f.proposalno)\n" +
                "on a.policyno = b.policyno\n" +
                "and a.registno = b.registno \n" +
                "left join cl_prplregist e\n" +
                "on a.registno = e.registno;");
        picczq.update("create index   registno on  temp11(registno); ");
        picczq.update("create temporary table temp12 \n" +
                "select a.*, b.checknature, c.checker1, c.checker2 \n" +
                "from temp11 a  left join cl_prplcheck b\n" +
                "on a.registno = b.registno\n" +
                "and b.validflag = '1'\n" +
                "left join  cl_prplchecktask c\n" +
                "on a.registno = c.registno\n" +
                "and c.validflag = '1';");
        picczq.update("create temporary table temp13 \t\n" +
                "select a.*,cast('' as nchar(20)) as licenseno0,cast('' as char(100)) as brandname0,cast('' as nchar(20)) as useyears0\n" +
                "from temp12 a;");
        picczq.update("create index registno on temp13(registno);");
        picczq.update("update temp13 set licenseno0 = (select max(licenseno) from cl_prplcitemcar\n" +
                "where temp13.registno = cl_prplcitemcar.registno);");
        picczq.update("update temp13 set brandname0 = (select max(brandname) from cl_prplcitemcar\n" +
                "where temp13.registno = cl_prplcitemcar.registno);");
        picczq.update("update temp13 set useyears0 = (select max(useyears) from cl_prplcitemcar\n" +
                "where temp13.registno = cl_prplcitemcar.registno);");

        //肇庆

        List<Map<String, Object>> result1 = picczq.queryForList("select * from temp13 order by enddeflossdate desc;");
        //车商

        List<Map<String, Object>> result2 = picczq.queryForList("select * from temp13\n" +
                "where comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342') \n" +
                "order by UNDERWRITEENDDATE desc;");
        //端州

        List<Map<String, Object>> result3 = picczq.queryForList("select * from temp13\n" +
                "where comcode in ('44129311','44120100','44120313','44129301','44129302','44129303','44129315','44129316','44129317','44129318','44129319','44129321','44129324','44129329','44129330','44129338')\n" +
                "order by enddeflossdate desc;");
        //鼎湖
        List<Map<String, Object>> result4 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like '441203%' and comcode <>'44120313'\n" +
                "order by enddeflossdate desc;");
        //四会
        List<Map<String, Object>> result5 = picczq.queryForList("select * from temp13\n" +
                "where comcode  like  '441284%'" +
                "or comcode in ('44129323','44129325')" +
                "order by enddeflossdate desc;");
        //广宁
        List<Map<String, Object>> result6 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441223%'\n" +
                "order by enddeflossdate desc;");
        //怀集

        List<Map<String, Object>> result7 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441224%'\n" +
                "order by enddeflossdate desc;");
        //高要

        List<Map<String, Object>> result8 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441283%'\n" +
                "order by enddeflossdate desc;");
        //德庆
        List<Map<String, Object>> result9 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441226%'\n" +
                "order by enddeflossdate desc;");
        //封开
        List<Map<String, Object>> result10 = picczq.queryForList("select * from temp13\n" +
                "where comcode   like  '441225%'\n" +
                "order by enddeflossdate desc;");

        picczq.update("drop table temp011_tmp;");
        picczq.update("drop table temp011;");
        picczq.update("drop table temp01;");
        picczq.update("drop table temp02;");
        picczq.update("drop table temp04;");
        picczq.update("drop table temp055;");
        picczq.update("drop table temp05;");
        picczq.update("drop table temp06;");
        picczq.update("drop table temp077;");
        picczq.update("drop table temp07;");
        picczq.update("drop table temp08;");
        picczq.update("drop table temp09;");
        picczq.update("drop table temp099;");
        picczq.update("drop table temp10;");
        picczq.update("drop table temp11;");
        picczq.update("drop table temp12;");
        picczq.update("drop table temp13;");

        //肇庆
        exportToExcel.hesun2(result1,"肇庆");

        //车商
        exportToExcel.hesun2(result2,"车商");

        //端州
        exportToExcel.hesun2(result3,"端州");

        //鼎湖
        exportToExcel.hesun2(result4,"鼎湖");

        //四会
        exportToExcel.hesun2(result5,"四会");

        //广宁
        exportToExcel.hesun2(result6,"广宁");

        //怀集
        exportToExcel.hesun2(result7,"怀集");

        //高要
        exportToExcel.hesun2(result8,"高要");

        //德庆
        exportToExcel.hesun2(result9,"德庆");

        //封开
        exportToExcel.hesun2(result10,"封开");

        return new ResultMap(1,"",true,"","");
    }
    //结案月
    public ResultMap monthjiean(){
        picczq.update("create temporary table jiean\n" +
                "select a2.comcode,a0.registno,b0.claimno,a2.lflag,b1.licenseno,b1.brandname,b0.policyno,b0.insuredname,b0.agentcode,b0.startdate,\n" +
                "b0.enddate,b0.businessnature,a0.damagedate,a0.damageaddress,a0.reportdate,c.endcasedate,a0.reportornumber,a2.repairfactorycode,a2.repairfactoryname,\n" +
                "a2.handlername,a2.deflossdate,a2.underwritename,truncate(a2.sumverilossfee,2),truncate(a2.sumlossfee,2),\n" +
                "cast('' as nchar(100)) as checker1,b.handlercode,\n" +
                "cast('' as nchar(100)) as handname,b.handler1code,\n" +
                "cast('' as nchar(100)) as hand1name,\n" +
                "cast('' as nchar(100)) as agentname,a2.underwriteenddate\n" +
                "from  cl_prplregist a0,cl_prpldeflossmain a2,cl_prplcmain b0,cl_prplcitemcar b1,cx_prpcmain b,cl_prplclaim c\n" +
                "where substring(a2.underwriteflag,1,1) in ('1','3') \n" +
                "and a2.validflag='1' \n" +
                "and c.endcasedate >=?\n" +
                "and c.endcasedate < ?\n" +
                "and c.casetype='2'\n" +
                "and c.endcasedate is not null\n" +
                "and a2.registno=b0.registno \n" +
                "and a2.riskcode=b0.riskcode \n" +
                "and b0.id=b1.prplcmainid \n" +
                "and a0.registno=c.registno\n" +
                "and a2.registno=a0.registno\n" +
                "and b0.policyno=b.policyno;",lastmonth,month);
        picczq.update("update jiean set handname = (select max(username) from utiihandler \n" +
                "where jiean.handlercode = utiihandler.usercode);");
        picczq.update("update jiean set checker1 = (select max(checker1) from cl_prplchecktask \n" +
                "where jiean.registno = cl_prplchecktask.registno);");
        picczq.update("update jiean set hand1name = (select max(username) from utiihandler \n" +
                "where jiean.handler1code = usercode);");
        picczq.update("update jiean set agentname = (select agentname from prpdagent \n" +
                "where jiean.agentcode = prpdagent.agentcode);");
        //肇庆

        List<Map<String, Object>> result1 = picczq.queryForList("select * from jiean order by endcasedate desc;");
        //车商

//
        List<Map<String, Object>> result2 = picczq.queryForList("select * from jiean\n" +
                "where comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342') \n" +
                "order by endcasedate desc;");
        //端州

//
        List<Map<String, Object>> result3 = picczq.queryForList("select * from jiean\n" +
                "where comcode in ('44129311','44120100','44120313','44129301','44129302','44129303','44129315','44129316','44129317','44129318','44129319','44129321','44129324','44129329','44129330','44129338')\n" +
                "order by endcasedate desc;");
        //鼎湖

//
        List<Map<String, Object>> result4 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like '441203%' and comcode <>'44120313'\n" +
                "order by endcasedate desc;");
        //四会


//
        List<Map<String, Object>> result5 = picczq.queryForList("select * from jiean\n" +
                "where comcode  like  '441284%'" +
                "or comcode in ('44129323','44129325')" +
                "order by endcasedate desc;");
        //广宁

        List<Map<String, Object>> result6 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441223%'\n" +
                "order by endcasedate desc;");
        //怀集

        List<Map<String, Object>> result7 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441224%'\n" +
                "order by endcasedate desc;");
        //高要

        List<Map<String, Object>> result8 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441283%'\n" +
                "order by endcasedate desc;");
        //德庆

        List<Map<String, Object>> result9 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441226%'\n" +
                "order by endcasedate desc;");
        //封开

        List<Map<String, Object>> result10 = picczq.queryForList("select * from jiean\n" +
                "where comcode   like  '441225%'\n" +
                "order by endcasedate desc;");
        picczq.update("drop table jiean");
        //肇庆
        exportToExcel.jiean2(result1,"肇庆");

        //车商
        exportToExcel.jiean2(result2,"车商");

        //端州
        exportToExcel.jiean2(result3,"端州");

        //鼎湖
        exportToExcel.jiean2(result4,"鼎湖");

        //四会
        exportToExcel.jiean2(result5,"四会");

        //广宁
        exportToExcel.jiean2(result6,"广宁");

        //怀集
        exportToExcel.jiean2(result7,"怀集");

        //高要
        exportToExcel.jiean2(result8,"高要");

        //德庆
        exportToExcel.jiean2(result9,"德庆");

        //封开
        exportToExcel.jiean2(result10,"封开");

        return null;
    }
    //保呗商城月
    public ResultMap bbscmonth(){
        picczq.update("create temporary table temp1(\n" +
                "select cast(null as char(120))comname,a.comcode,cast(null as char(120))comcodename,a.proposalno,a.policyno,a.sumpremium,\n" +
                "a.sumtaxpremium,a.sumnetpremium,a.agentcode,cast(null as char(120)) agentname,b.clausetype,\n" +
                "b.carkindcode,cast(null as char(120))carkindname,a.startdate,a.enddate,b.licenseno,a.businessnature,cast(null as char(120))businessnaturename,\n" +
                "a.operatedate,a.riskcode,b.monopolycode,b.monopolyname,a.handler1code,cast(null as char(120))handler1name,\n" +
                "a.handlercode,cast(null as char(120))handlername,cast(null as datetime) policynotime,a.projectcode,\n" +
                "c.clausecode,c.clausename,c.clauses,b.newcarflag,cast(null as char(22))lastpolicyno,\n" +
                "b.usenaturecode,cast(null as char(20)) usenaturename,cast(null as char(20))xinxuzhuan,\n" +
                "cast(null as char(120)) insuredname,cast(null as char(120)) bbrinsuredname,cast(null as char(4))insuredtype,cast(0.0 as decimal(65,4))jifen\n" +
                "from cx_prpcmain a,cx_prpcitem_car b,cx_prpcengage c\n" +
                "where c.clausecode in ('997468','997469','997470','997471','997472','9908983','9908985','9908987','9908988','9908989','9908991','9908992','000121')\n" +
                "and a.operatedate >= ?\n" +
                "and a.operatedate < ?\n" +
                "and a.policyno is not null\n" +
                "and a.policyno != ''\n" +
                "and a.proposalno = b.proposalno\n" +
                "and a.proposalno = c.proposalno\n" +
                "); ",lastmonth,month);
        picczq.update("create index proposalno on temp1(proposalno);");
        picczq.update("UPDATE temp1\n" +
                "    SET jifen = CASE temp1.clausecode \n" +
                "        when '997468' then '200'\n" +
                "        when '997469' then '400'\n" +
                "        when '997470' then '600'\n" +
                "        when '997471' then '800'\n" +
                "        when '997472' then '1000'\n" +
                "        when '9908983' then floor(sumnetpremium * 0.05)\n" +
                "        when '9908985' then floor(sumnetpremium * 0.10)\n" +
                "        when '9908987' then floor(sumnetpremium * 0.15)\n" +
                "        when '9908988' then floor(sumnetpremium * 0.20)\n" +
                "        when '9908989' then floor(sumnetpremium * 0.25)\n" +
                "        when '9908991' then floor(sumnetpremium * 0.30)\n" +
                "        when '9908992' then floor(sumnetpremium * 0.35)\n" +
                "    END\n" +
                "WHERE temp1.clausecode IN ('997468','997469','997470','997471','997472','9908983','9908985','9908987','9908988','9908989','9908991','9908992');");

        picczq.update("update temp1 set comname = '车商'\n" +
                "        where  comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342');");
        picczq.update("update temp1 set comname = '端州'\n" +
                "        where comcode in ('44120100','44129301','44129302','44129303','44129311','44129315','44129317','44129318','44129319',\n" +
                "              '44129321','44129329','44129330','44129338','44129343','44129324','44120313');");
        picczq.update("update temp1 set comname = '鼎湖'\n" +
                "        where comcode like '441203%' and comcode <>'44120313';");
        picczq.update("update temp1 set comname = '广宁'\n" +
                "        where comcode like '441223%';");
        picczq.update("update temp1 set comname = '怀集'\n" +
                "        where comcode like '441224%';");
        picczq.update("update temp1 set comname = '封开'\n" +
                "        where comcode like '441225%';");
        picczq.update("update temp1 set comname = '德庆'\n" +
                "        where comcode like '441226%';");
        picczq.update("update temp1 set comname = '高要'\n" +
                "        where comcode like '441283%';");
        picczq.update("update temp1 set comname = '四会'\n" +
                "        where comcode like '441284%';");
        picczq.update("update temp1 set comname = '大旺'\n" +
                "        where comcode in ('44129323','44129325');");
        picczq.update("update temp1 set comname = '电销'\n" +
                "        where comcode in ('44129400');");
        picczq.update("create temporary table temp2(\n" +
                "select proposalno,policyno,oldpolicyno\n" +
                "from cx_prpcrenewal \n" +
                "where proposalno in (select proposalno from temp1));");
        picczq.update("update temp1 set xinxuzhuan = '转保';");
        picczq.update("update temp1 set temp1.xinxuzhuan = '新保' \n" +
                "where newcarflag = 1;");
        picczq.update("update temp1,temp2 set temp1.xinxuzhuan = '续保' \n" +
                "where temp1.proposalno = temp2.proposalno;");
        picczq.update("update temp1,temp2 set  lastpolicyno = oldpolicyno\n" +
                "where temp1.proposalno  = temp2.proposalno;");
        picczq.update("update temp1,prpdcompany set comcodename = comcname\n" +
                "where temp1.comcode = prpdcompany.comcode;");
        picczq.update("update temp1,prpdagent set temp1.agentname = prpdagent.agentname\n" +
                "where temp1.agentcode = prpdagent.agentcode;");
        picczq.update("update temp1\n" +
                "    set carkindname = case temp1.carkindcode \n" +
                "        when 'a01' then '客车'\n" +
                "        when 'b01' then '货车'\n" +
                "        when 'b02' then '半挂牵引车'\n" +
                "        when 'b11' then '三轮汽车'\n" +
                "        when 'b12' then '低速货车'\n" +
                "        when 'b13' then '客货两用车'\n" +
                "        when 'b21' then '自卸货车'\n" +
                "        when 'b91' then '货车挂车'\n" +
                "        when 'c01' then '油罐车'\n" +
                "        when 'c02' then '气罐车'\n" +
                "        when 'c03' then '液罐车'\n" +
                "        when 'c04' then '冷藏车'\n" +
                "        when 'c11' then '罐车挂车'\n" +
                "        when 'c20' then '推土车'\n" +
                "        when 'c22' then '清障车'\n" +
                "        when 'c23' then '清扫车'\n" +
                "        when 'c24' then '清洁车'\n" +
                "        when 'c25' then '起重车'\n" +
                "        when 'c26' then '装卸车'\n" +
                "        when 'c27' then '升降车'\n" +
                "        when 'c28' then '混凝土搅拌车'\n" +
                "        when 'c29' then '挖掘车'\n" +
                "        when 'c30' then '专业拖车'\n" +
                "        when 'c31' then '特种车二挂车'\n" +
                "        when 'c39' then '特种车二类其它'\n" +
                "        when 'c41' then '电视转播车'\n" +
                "        when 'c42' then '消防车'\n" +
                "        when 'c43' then '医疗车'\n" +
                "        when 'c44' then '油汽田操作用车'\n" +
                "        when 'c45' then '压路车'\n" +
                "        when 'c46' then '矿山车'\n" +
                "        when 'c47' then '运钞车'\n" +
                "        when 'c48' then '救护车'\n" +
                "        when 'c49' then '监测车'\n" +
                "        when 'c50' then '雷达车'\n" +
                "        when 'c51' then 'x光检查车'\n" +
                "        when 'c52' then '电信抢修车/电信工程车'\n" +
                "        when 'c53' then '电力抢修车/电力工程车'\n" +
                "        when 'c54' then '专业净水车'\n" +
                "        when 'c55' then '保温车'\n" +
                "        when 'c56' then '邮电车'\n" +
                "        when 'c57' then '警用特种车'\n" +
                "        when 'c58' then '混凝土泵车'\n" +
                "        when 'c61' then '特种车三类挂车'\n" +
                "        when 'c69' then '特种车三类其它'\n" +
                "        when 'c90' then '集装箱拖头'\n" +
                "        when 'd01' then '摩托车'\n" +
                "        when 'd02' then '正三轮摩托车'\n" +
                "        when 'd03' then '侧三轮摩托车'\n" +
                "        when 'e01' then '拖拉机'\n" +
                "        when 'e11' then '联合收割机'\n" +
                "        when 'e12' then '变形拖拉机/其它'\n" +
                "        when 'z99' then '其它车辆'\n" +
                "    end\n" +
                "where temp1.carkindcode in ('a01','b01','b02','b11','b12','b13','b21','b91','c01','c02','c03','c04','c11','c20',\n" +
                "'c22','c23','c24','c25','c26','c27','c28','c29','c30','c31','c39','c41','c42','c43','c44','c45','c46','c47','c48',\n" +
                "'c49','c50','c51','c52','c53','c54','c55','c56','c57','c58','c61','c69','c90','d01','d02','d03','e01','e11','e12','z99'\n" +
                ");");
        picczq.update("update temp1\n" +
                "    set businessnaturename = case temp1.businessnature\n" +
                "        when '1' then '个人代理业务'\n" +
                "        when '0' then '传统直销业务'\n" +
                "        when '2' then '专业代理业务'\n" +
                "        when '3' then '兼业代理业务'\n" +
                "        when '4' then '经纪业务'\n" +
                "        when '53' then '网上业务'\n" +
                "    end\n" +
                "where temp1.businessnature in ('1','0', '2','3','4','53');");
        picczq.update("update temp1,utiihandler set handler1name = username\n" +
                "where utiihandler.usercode = temp1.handler1code;");
        picczq.update("update temp1,utiihandler set handlername = username\n" +
                "where utiihandler.usercode = temp1.handlercode;");
        picczq.update("update temp1,prptime set temp1.policynotime = prptime.operatetimeforhis\n" +
                "where temp1.policyno = prptime.certino\n" +
                "and prptime.updatetype = 'c01';");
        picczq.update("update temp1\n" +
                "    set usenaturename = case temp1.usenaturecode \n" +
                "        when '000' then '不区分营业非营业'\n" +
                "        when '111' then '出租、租赁'\n" +
                "        when '112' then '城市公交'\n" +
                "        when '113' then '公路客运'\n" +
                "        when '114' then '旅游客运'\n" +
                "        when '120' then '营业货车'\n" +
                "        when '121' then '营业挂车'\n" +
                "        when '180' then '运输型拖拉机'\n" +
                "        when '190' then '其它营业车辆'\n" +
                "        when '211' then '家庭自用汽车'\n" +
                "        when '212' then '非营业企业客车'\n" +
                "        when '213' then '非营业机关、事业团体客车'\n" +
                "        when '220' then '非营业货车'\n" +
                "        when '221' then '非营业挂车'\n" +
                "        when '280' then '兼用型拖拉机'\n" +
                "        when '290' then '其它非营业车辆'\n" +
                "    end\n" +
                "where temp1.usenaturecode in ('000','111','112','113','114','120','121','180','190','211','212','213','220','221','280','290');");
        picczq.update("update temp1,cx_prpcinsured set temp1.insuredtype = '个人'\n" +
                "where cx_prpcinsured.proposalno = temp1.proposalno\n" +
                "and cx_prpcinsured.insuredtype = 1;");
        picczq.update("update temp1,cx_prpcinsured set temp1.insuredtype = '团体'\n" +
                "where cx_prpcinsured.proposalno = temp1.proposalno\n" +
                "and cx_prpcinsured.insuredtype = 2;");
        picczq.update("update temp1 a,cx_prpcinsured b set a.insuredname = b.insuredname\n" +
                "where a.proposalno = b.proposalno\n" +
                "and substring(b.insuredflag,1,1) = 1;");
        picczq.update("update temp1 a,cx_prpcinsured b set a.bbrinsuredname = b.insuredname\n" +
                "where a.proposalno = b.proposalno\n" +
                "and substring(b.insuredflag,2,1) = 1;");
        picczq.update("create temporary table temp3(\n" +
                "select a.proposalno,a.serialno,cast(null as char(120))mainsellfeerate,sellfeerate,\n" +
                "cast(null as char(120))summainsellfeerate,cast(null as char(120))sumsellfeerate,\n" +
                "cast(null as char(120))maxcostrate,a.mainflag \n" +
                "from cx_prpcseller a,temp1 b\n" +
                "where b.proposalno = a.proposalno\n" +
                ");");
        picczq.update("create temporary table temp4(\n" +
                "select sum(sellfeerate)sellfeerate,proposalno,mainflag\n" +
                "from temp3\n" +
                "group by proposalno,mainflag\n" +
                ");\n");
        picczq.update("update temp3 a,temp4 b\n" +
                "set mainsellfeerate = b.sellfeerate,summainsellfeerate = b.sellfeerate\n" +
                "where  b.mainflag = 1\n" +
                "and a.proposalno = b.proposalno;");
        picczq.update("update temp3 a,temp4 b\n" +
                "set sumsellfeerate = b.sellfeerate\n" +
                "where  b.mainflag = 0\n" +
                "and a.proposalno = b.proposalno;");
        picczq.update("create temporary table temp5(\n" +
                "select cast(null as char(120))proposalno,a.policyno,max(a.maxcostrate)maxcostrate from scmsdocfeedetail a,temp1 b\n" +
                "where a.certino = b.policyno\n" +
                "and a.certitype = 'p'\n" +
                "group by a.policyno\n" +
                ");");
        picczq.update("update temp5 a,temp1 b set a.proposalno = b.proposalno\n" +
                "where a.policyno = b.policyno;");
        picczq.update("update temp3 a,temp5 b set a.maxcostrate = b.maxcostrate\n" +
                "where a.proposalno = b.proposalno;");

        List<Map<String, Object>> result  = picczq.queryForList("select a.comname,a.comcodename,a.policyno,a.sumpremium,\n" +
                "a.sumtaxpremium,a.sumnetpremium,a.agentname,a.clausetype,\n" +
                "a.carkindname,a.startdate,a.enddate,a.licenseno,a.businessnaturename,\n" +
                "a.operatedate,a.riskcode,a.monopolycode,a.monopolyname,a.handler1name,\n" +
                "a.handlername,a.policynotime,a.projectcode,a.insuredname,a.bbrinsuredname,\n" +
                "a.xinxuzhuan,b.mainsellfeerate,b.sellfeerate,b.summainsellfeerate,\n" +
                "b.sumsellfeerate,b.maxcostrate,a.lastpolicyno,a.usenaturename,\n" +
                "a.insuredtype,a.clausecode,a.clausename,a.clauses\n" +
                "from temp1 a,temp3 b\n" +
                "where a.proposalno = b.proposalno;");

        List<Map<String, Object>> result2 = picczq.queryForList("select sum(sumnetpremium)sumnetpremium,clausecode,clausename,sum(jifen) sumjifen,count(*)count from temp1\n" +
                "group by clausecode");
        List<Map<String, Object>> total = picczq.queryForList("select sum(jifen)total from temp1\n");
        picczq.update("drop table temp1");
        picczq.update("drop table temp2");
        picczq.update("drop table temp3");
        picczq.update("drop table temp4");
        picczq.update("drop table temp5");


        exportToExcel.bbscmonth(result);
        exportToExcel.bbsc2month(result2,total);

        return  new ResultMap(1,"",true,"","");

    }
    //交叉
    public ResultMap jiaocha(){

        //车险交叉销售
        picczq.update("create temporary table fc1\n" +
                "select a.comcode,a.riskcode,a.proposalno,a.policyno,a.contractno,a.projectcode,a.sumamount,a.sumpremium,a.sumpremium-a.sumtaxpremium as jbf,a.sumtaxpremium,\n" +
                "a.startdate,a.enddate,a.operatedate,a.agentcode,a.handlercode,a.handler1code,\n" +
                "cast('' as nchar(100)) as zqdfl,\n" +
                "cast('' as nchar(100)) as zqdfy,\n" +
                "cast('' as nchar(100)) as fqdfl,\n" +
                "cast('' as nchar(100)) as fqdfy\n" +
                "from cx_prpcmain a\t\n" +
                "where a.underwriteflag in ('1','3')\n" +
                "and a.agentcode  in ('44003f100013','44013j200001','44003j200022')\n" +
                "and a.operatedate >=?  and a.operatedate <?;",lastmonth,month);
        picczq.update("create  index policyno on fc1 (policyno);");
        picczq.update("update fc1 set zqdfl = (select sum(sellfeerate) from cx_prpcseller \n" +
                "where fc1.policyno = cx_prpcseller.policyno \n" +
                "and cx_prpcseller.mainflag = '1');");
        picczq.update("update fc1 set zqdfy = (select sum(sellfee) from cx_prpcseller \n" +
                "where fc1.policyno = cx_prpcseller.policyno \n" +
                "and cx_prpcseller.mainflag = '1');");
        picczq.update("update fc1 set fqdfl = (select sum(sellfeerate) from cx_prpcseller \n" +
                "where fc1.policyno = cx_prpcseller.policyno \n" +
                "and cx_prpcseller.mainflag = '0');");
        picczq.update("update fc1 set fqdfy = (select sum(sellfee) from cx_prpcseller \n" +
                "where fc1.policyno = cx_prpcseller.policyno \n" +
                "and cx_prpcseller.mainflag = '0');");
        List<Map<String, Object>> result1 = picczq.queryForList("select * from fc1;");
        picczq.update("drop table fc1;");
        exportToExcel.cxjiaocha(result1);


        //非车险交叉销售
        picczq.update("create temporary table fc1\n" +
                "select a.comcode,a.riskcode,a.proposalno,a.policyno,a.contractno,a.projectcode,a.sumamount,a.sumpremium,a.sumpremium-a.sumtaxfee as jbf,a.sumtaxfee,\n" +
                "a.startdate,a.enddate,a.operatedate,a.agentcode,a.handlercode,a.handler1code,\n" +
                "cast('' as nchar(100)) as zqdfl,\n" +
                "cast('' as nchar(100)) as zqdfy,\n" +
                "cast('' as nchar(100)) as fqdfl,\n" +
                "cast('' as nchar(100)) as fqdfy\n" +
                "from fc_prpcmain a\t\n" +
                "where a.underwriteflag in ('1','3')\n" +
                "and a.agentcode  in ('44003f100013','44013j200001','44003j200022')\n" +
                "and a.operatedate >=?  and a.operatedate <?;",lastmonth,month);
        picczq.update("create  index policyno on fc1 (policyno);");
        picczq.update("update fc1 set zqdfl = (select sum(sellfeerate) from fc_prpcseller \n" +
                "where fc1.policyno = fc_prpcseller.policyno \n" +
                "and fc_prpcseller.mainflag = '1');");
        picczq.update("update fc1 set zqdfy = (select sum(sellfee) from fc_prpcseller \n" +
                "where fc1.policyno = fc_prpcseller.policyno \n" +
                "and fc_prpcseller.mainflag = '1');");
        picczq.update("update fc1 set fqdfl = (select sum(sellfeerate) from fc_prpcseller \n" +
                "where fc1.policyno = fc_prpcseller.policyno \n" +
                "and fc_prpcseller.mainflag = '0');");
        picczq.update("update fc1 set fqdfy = (select sum(sellfee) from fc_prpcseller \n" +
                "where fc1.policyno = fc_prpcseller.policyno \n" +
                "and fc_prpcseller.mainflag = '0');");
        List<Map<String, Object>> result2 = picczq.queryForList("select * from fc1;");
        picczq.update("drop table fc1;");
        exportToExcel.fcjiaocha(result2);
        return new ResultMap(1,"",true,"","");
    }
    //车商口径
    public ResultMap cheshangkoujing(){
        picczq.update("create temporary table temp1 (\n" +
                "select proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode,cast(null as char(20)) xuzhuanbao\n" +
                "from cx_prpcmain\n" +
                "where startdate >= ? and startdate < ?\n" +
                "and comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342','44129346')\n" +
                "and policyno != ''\n" +
                "and policyno is not null);",lastmonth,month);
        picczq.update("create index proposalno on temp1(proposalno);");
        picczq.update("create index policyno on temp1(policyno);");
        picczq.update("create temporary table temp2(\n" +
                "select proposalno,clausetype,licenseno,frameno,engineno,enrolldate,carkindcode,\n" +
                "monopolycode,monopolyname,newcarflag\n" +
                "from cx_prpcitem_car \n" +
                "where proposalno in (select proposalno from temp1));");
        picczq.update("create index temp2_proposalno on temp2(proposalno);");
        picczq.update("create temporary table temp3(\n" +
                "select proposalno,policyno,oldpolicyno\n" +
                "from cx_prpcrenewal \n" +
                "where proposalno in (select proposalno from temp1));");
        picczq.update("update temp1 set xuzhuanbao = '转保';");
        picczq.update("update temp1,temp2 set temp1.xuzhuanbao = '新保' \n" +
                "where temp1.proposalno = temp2.proposalno\n" +
                "and temp2.newcarflag = 1;");
        picczq.update("update temp1,temp3 set temp1.xuzhuanbao = '续保' \n" +
                "where temp1.proposalno = temp3.proposalno;");
        List<Map<String, Object>> result = picczq.queryForList("select a.proposalno,a.policyno,a.riskcode,a.operatedate,a.startdate,\n" +
                "a.enddate,a.sumamount,a.sumpremium,a.sumnetpremium,a.handlercode,a.comcode,a.handler1code,\n" +
                "a.agentcode,a.xuzhuanbao,b.clauseType,b.LicenseNo,b.FrameNo,b.EngineNo,b.EnrollDate,\n" +
                "b.carKindCode,b.MonopolyCode,b.MonopolyName,b.NewCarFlag\n" +
                "from temp1 a,temp2 b\n" +
                "where b.proposalno = a.proposalno order by startdate,comcode;");

        picczq.update("drop table temp1;");
        picczq.update("drop table temp2;");
        picczq.update("drop table temp3;");
        exportToExcel.cskj(result);
        return new ResultMap(1,"",true,"","");
    }
    //车险起保
    public ResultMap chexianqibao(){
        picczq.update("create temporary table temp1 (\n" +
                "select proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode,cast('转保' as char(20))xinxuzhuan\n" +
                "from cx_prpcmain\n" +
                "where (startdate >=? and startdate<? and policyno like 'p%')\n" +
                "or (startdate >=? and startdate<? and policyno like 'p%'));",lastmonth,month,lastyearlastmonth,lastyearmonth);
        picczq.update("create index proposalno on temp1(proposalno);\n");
        picczq.update("create index policyno on temp1(policyno);\n");
        picczq.update("create temporary table temp2(\n" +
                "select proposalno,clausetype,licenseno,frameno,newcarflag\n" +
                "from cx_prpcitem_car \n" +
                "where proposalno in (select proposalno from temp1));");
        picczq.update("create index temp2_proposalno on temp2(proposalno);\n");


        picczq.update("create temporary table temp22(\n" +
                "select proposalno,policyno,oldpolicyno\n" +
                "from cx_prpcrenewal \n" +
                "where proposalno in (select proposalno from temp1));");

        picczq.update("update temp1,temp2 set xinxuzhuan = '新保'\n" +
                "where temp1.proposalno = temp2.proposalno\n" +
                "and temp2.newcarflag = 1;");
        picczq.update("update temp1,temp22 set temp1.xinxuzhuan = '续保' \n" +
                "where temp1.proposalno = temp22.proposalno;");


        List<Map<String, Object>> result = picczq.queryForList("select a.proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode,\n" +
                "clauseType,LicenseNo,FrameNo,a.xinxuzhuan from temp1 a,temp2 b \n" +
                "where b.proposalno = a.proposalno order by startdate;");
        picczq.update("drop table temp1;");
        picczq.update("drop table temp2;");
        exportToExcel.cxqb(result);

        return new ResultMap(1,"",true,"","");
    }
    //车险起保week
    public ResultMap chexianqibaoweek(){
        picczq.update("create temporary table temp1 (\n" +
                "select proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode,cast('转保' as char(20))xinxuzhuan\n" +
                "from cx_prpcmain\n" +
                "where (startdate >=? and startdate<? and policyno like 'p%')\n" +
                "or (startdate >=? and startdate<? and policyno like 'p%'));",lastmonth,month,lastyearlastmonth,lastyearmonth);
        picczq.update("create index proposalno on temp1(proposalno);\n");
        picczq.update("create index policyno on temp1(policyno);\n");
        picczq.update("create temporary table temp2(\n" +
                "select proposalno,clausetype,licenseno,frameno,newcarflag\n" +
                "from cx_prpcitem_car \n" +
                "where proposalno in (select proposalno from temp1));");
        picczq.update("create index temp2_proposalno on temp2(proposalno);\n");


        picczq.update("create temporary table temp22(\n" +
                "select proposalno,policyno,oldpolicyno\n" +
                "from cx_prpcrenewal \n" +
                "where proposalno in (select proposalno from temp1));");

        picczq.update("update temp1,temp2 set xinxuzhuan = '新保'\n" +
                "where temp1.proposalno = temp2.proposalno\n" +
                "and temp2.newcarflag = 1;");
        picczq.update("update temp1,temp22 set temp1.xinxuzhuan = '续保' \n" +
                "where temp1.proposalno = temp22.proposalno;");


        List<Map<String, Object>> result = picczq.queryForList("select a.proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode,\n" +
                "clauseType,LicenseNo,FrameNo,a.xinxuzhuan from temp1 a,temp2 b \n" +
                "where b.proposalno = a.proposalno order by startdate;");
        picczq.update("drop table temp1;");
        picczq.update("drop table temp2;");
        exportToExcel.cxqb(result);

        return new ResultMap(1,"",true,"","");
    }
    //车险签单
    public ResultMap chexianqiandan(){
        picczq.update("create temporary table temp1 (\n" +
                "select proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode\n" +
                "from cx_prpcmain\n" +
                "where operatedate>= ? and operatedate< ?\n" +
                "and agentcode in ('44003f100013','44013j200001','44003j200022')\n" +
                "and policyno != ''\n" +
                "and policyno is not null);",lastmonth,month);
        picczq.update("create index proposalno on temp1(proposalno);");
        picczq.update("create index policyno on temp1(policyno);");
        picczq.update("create temporary table temp2(\n" +
                "select proposalno,clausetype,licenseno,frameno\n" +
                "from cx_prpcitem_car \n" +
                "where proposalno in (select proposalno from temp1));");
        picczq.update("create index temp2_proposalno on temp2(proposalno);");
        List<Map<String, Object>> result = picczq.queryForList("select a.proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode,\n" +
                "clauseType,LicenseNo,FrameNo from temp1 a,temp2 b\n" +
                "where b.proposalno = a.proposalno;");
        picczq.update("drop table temp1");
        picczq.update("drop table temp2");
        exportToExcel.cxqd(result);

        return new ResultMap(1,"",true,"","");
    }
    //非车起保
    public ResultMap feicheqibao(){
        picczq.update("create temporary table temp1 (\n" +
                "select proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode\n" +
                "from fc_prpcmain\n" +
                "where (startdate>=?  and startdate<? and policyno like '%p%' and riskcode not in ('hu3','mse','msd','muw','iay','ian','iau','hu2','iln','iji','ig6','mpl','iul'))\n" +
                "or (startdate>=?  and startdate<? and policyno like '%p%' and riskcode not in ('hu3','mse','msd','muw','iay','ian','iau','hu2','iln','iji','ig6','mpl','iul'))\n" +
                ");",lastmonth,month,lastyearlastmonth,lastyearmonth);
        List<Map<String, Object>> result = picczq.queryForList("select proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode from temp1 order by startdate;");
        picczq.update("drop table temp1;");
        exportToExcel.fcqb(result);

        return new ResultMap(1,"",true,"","");
    }
    //非车签单
    public ResultMap feicheqiandan(){

        picczq.update("create temporary table temp1 (\n" +
                "select proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode\n" +
                "from fc_prpcmain\n" +
                "where (operatedate>=?  and operatedate<? )\n" +
                "and agentcode in ('44003f100013','44013j200001','44003j200022')\n" +
                "and policyno != ''\n" +
                "and policyno is not null);",lastmonth,month);
        List<Map<String, Object>> result = picczq.queryForList("select proposalno,policyno,riskcode,operatedate,startdate,\n" +
                "enddate,sumamount,sumpremium,sumnetpremium,handlercode,comcode,handler1code,agentcode from temp1 order by operatedate,agentcode;");
        picczq.update("drop table temp1;");
        exportToExcel.fcqd(result);


        return new ResultMap(1,"",true,"","");
    }
    //车船税
    public ResultMap chechuanshui(){
        List<Map<String, Object>> result1 = picczq.queryForList("select  a.taxpayercode,a.dutypaidproofno,a.taxprintproofno,a.licenseno,a.model,a.carkindcode,a.taxpayeridentno,a.thispaytax,a.prepaytax,a.delaypaytax,a.sumpaytax,\n" +
                "b.policyno,b.startdate,b.enddate,d.deskdate,b.comcode,a.taxpayername,a.taxpayernumber,b.handlercode,a.taxtype,a.frameno,a.engineno,a.enrolldate,a.carloteququality,c.exhaustscale,\n" +
                "d.kindcode,d.sffcomcode,d.mainamount,d.rateamount,b.agentcode,c.monopolycode,c.monopolyname,c.modelcode,c.brandname,c.usenaturecode,c.carkindcode,\n" +
                "a.taxabatereason,c.seatcount\n" +
                "from cx_prpccarshiptax a,cx_prpcmain b,cx_prpcitem_car c,sff_sffdocdetail d\n" +
                "where a.proposalno=b.proposalno\n" +
                "and b.proposalno=c.proposalno\n" +
                "and b.policyno=d.policyno\n" +
                "and b.underwriteflag in ('1','3')\n" +
                "and d.kindcode in ('rm9','rm7','pm9','pm7','pm8')\n" +
                "and d.deskdate >=? and d.deskdate <?;",lastmonth,month);

        exportToExcel.chechuanshui(result1);

        return new ResultMap(1,"",true,"","");
    }
    //出单统计量
    public ResultMap cdtjl(){
        //车商车险出单
        List<Map<String, Object>> result1 = picczq.queryForList(" select '车险保单',a.comcode,a.policyno,a.ProposalNo,a.RiskCode,a.inputtime,a.StartDate,a.SumPremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.BusinessNature<>'5'\n" +
                "and a.OperatorCode in ('A441200024 ','A441200026','A441200083','A441200084','A441200092','A441200100','A441200101','A441200102')\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.RiskCode in ('DZA','DAA')\n" +
                "group by a.operatorcode,b.licenseno",lastmonth,month);
        //三版车险出单
        List<Map<String, Object>> result2 = picczq.queryForList("select '车险保单',a.comcode,a.policyno,a.ProposalNo,a.RiskCode,a.inputtime,a.StartDate,a.SumPremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.BusinessNature<>'5'\n" +
                "and a.OperatorCode in ('11100901','11100908','11100930','11100935','15204102','4412001028','4412001035','4412001036','4412031015','4412251010','4412255009','4412261004','4412261005','4412261007','4412845007','4412831021')\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.RiskCode in ('DZA','DAA')\n" +
                "group by b.licenseno\n" +
                "union \n" +
                "select '车险保单', a.comcode,a.policyno,a.ProposalNo,a.RiskCode,a.inputtime,a.StartDate,a.SumPremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.BusinessNature<>'5'\n" +
                "and a.OperatorCode in ('4412831025')\n" +
                "and a.comcode = '44128314'\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.RiskCode in ('DZA','DAA')\n" +
                "group by b.licenseno\n" +
                "union \n" +
                "select '车险保单', a.comcode,a.policyno,a.ProposalNo,a.RiskCode,a.inputtime,a.StartDate,a.SumPremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.BusinessNature<>'5'\n" +
                "and a.OperatorCode in ('4412835047')\n" +
                "and a.comcode = '44128318'\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.RiskCode in ('DZA','DAA')\n" +
                "group by b.licenseno",lastmonth,month,lastmonth,month,lastmonth,month);
        //非车商个代出单
        List<Map<String, Object>> result3 = picczq.queryForList("select '车险保单',a.comcode,a.policyno,a.ProposalNo,a.RiskCode,a.inputtime,a.StartDate,a.SumPremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.BusinessNature<>'5'\n" +
                "and a.OperatorCode in ('15188181','44245857','44498880','84402863','A441200085','A441200090','A441200091','A441200093','A441200094','A441200095','A441200096')\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.RiskCode in ('DZA','DAA')\n" +
                "group by b.licenseno",lastmonth,month);

        //车商保单批单
        List<Map<String, Object>> result4 = picczq.queryForList(" select '车险保单',a.comcode,a.policyno,a.ProposalNo,a.RiskCode,a.inputtime,a.StartDate,a.SumPremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.BusinessNature<>'5'\n" +
                "and a.OperatorCode in ('A441200024 ','A441200026','A441200083','A441200084','A441200092','A441200100','A441200101','A441200102')\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.RiskCode in ('DZA','DAA','DZB')\n" +
                "and a.policyno like '%P%'\n" +
                "union \n" +
                "select '非车保单',comcode,policyno,ProposalNo,RiskCode,inputtime,StartDate,SumPremium,OperatorCode,'非车险' as  licenseno\n" +
                "from fc_prpcmain \n" +
                "where underwriteflag in ('1','3')\n" +
                "and OperatorCode in ('A441200024 ','A441200026','A441200083','A441200084','A441200092','A441200100','A441200101','A441200102')\n" +
                "and (proposalno like 'T%' or proposalno is null)\n" +
                "and BusinessNature<>'5'\n" +
                "and inputtime >=?\t\n" +
                "and inputtime <?\t\t\n" +
                "union \n" +
                "select '车险批单',comcode,policyno,EndorseNo,riskcode,inputtime,EndorDate,0 as SumPremium,OperatorCode,'车P' as  licenseno\n" +
                "from cx_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and OperatorCode in ('A441200024 ','A441200026','A441200083','A441200084','A441200092','A441200100','A441200101','A441200102')\n" +
                "and SUBSTRING(AgentCode,5,1)<>'5'\n" +
                "and EndorDate >=?\n" +
                "and EndorDate <?\t\t\t\t\n" +
                "union \n" +
                "select '非车批单',comcode,policyno,EndorseNo,riskcode,inputtime,EndorDate,0 as SumPremium,OperatorCode,'非车P' as  licenseno\n" +
                "from fc_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and OperatorCode in ('A441200024 ','A441200026','A441200083','A441200084','A441200092','A441200100','A441200101','A441200102')\n" +
                "and SUBSTRING(AgentCode,5,1)<>'5'\n" +
                "and EndorDate >=?\n" +
                "and EndorDate <?;",lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month);
        //三版保单批单
        List<Map<String, Object>> result5 = picczq.queryForList("  select '车险保单',a.comcode,a.policyno,a.proposalno,a.riskcode,a.inputtime,a.startdate,a.sumpremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.businessnature<>'5'\n" +
                "and a.operatorcode in ('11100901','11100908','11100930','11100935','15204102','4412001028','4412001035','4412001036','4412031015','4412251010','4412255009','4412261004','4412261005','4412261007','4412845007','4412831021')\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.riskcode in ('dza','daa')\n" +
                "and a.policyno like '%p%'\n" +
                "union\n" +
                "select '车险保单',a.comcode,a.policyno,a.proposalno,a.riskcode,a.inputtime,a.startdate,a.sumpremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.businessnature<>'5'\n" +
                "and a.operatorcode in ('4412831025')\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.riskcode in ('DZA','DAA','DZB')\n" +
                "and a.policyno like '%p%'\n" +
                "and a.comcode = '44128314'\n" +
                "union  \n" +
                "select '车险保单',a.comcode,a.policyno,a.proposalno,a.riskcode,a.inputtime,a.startdate,a.sumpremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.businessnature<>'5'\n" +
                "and a.operatorcode in ('4412835047')\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.riskcode in ('dza','daa')\n" +
                "and a.policyno like '%p%'\n" +
                "and a.comcode = '44128318'\n" +
                "group by a.operatorcode,b.licenseno\n" +
                "union\n" +
                "select '非车保单',comcode,policyno,proposalno,riskcode,inputtime,startdate,sumpremium,operatorcode,'非车险' as  licenseno\n" +
                "from fc_prpcmain \n" +
                "where underwriteflag in ('1','3')\n" +
                "and operatorcode in ('11100901','11100908','11100930','11100935','15204102','4412001028','4412001035','4412001036','4412031015','4412251010','4412255009','4412261004','4412261005','4412261007','4412845007','4412831021')\n" +
                "and (proposalno like 't%' or proposalno is null)\n" +
                "and businessnature<>'5'\n" +
                "and inputtime >=?\t\n" +
                "and inputtime <?\t\t\n" +
                "union \n" +
                "select '非车保单',comcode,policyno,proposalno,riskcode,inputtime,startdate,sumpremium,operatorcode,'非车险' as  licenseno\n" +
                "from fc_prpcmain \n" +
                "where underwriteflag in ('1','3')\n" +
                "and operatorcode in ('4412835047')\n" +
                "and (proposalno like 't%' or proposalno is null)\n" +
                "and businessnature<>'5'\n" +
                "and comcode = '44128318'\n" +
                "and inputtime >=?\t\n" +
                "and inputtime <?\n" +
                "union\n" +
                "select '非车保单',comcode,policyno,proposalno,riskcode,inputtime,startdate,sumpremium,operatorcode,'非车险' as  licenseno\n" +
                "from fc_prpcmain \n" +
                "where underwriteflag in ('1','3')\n" +
                "and operatorcode in ('4412831025')\n" +
                "and (proposalno like 't%' or proposalno is null)\n" +
                "and businessnature<>'5'\n" +
                "and comcode = '44128314'\n" +
                "and inputtime >=?\t\n" +
                "and inputtime <?\n" +
                "\n" +
                "union\n" +
                "select '车险批单',comcode,policyno,endorseno,riskcode,inputtime,endordate,0 as sumpremium,operatorcode,'车p' as  licenseno\n" +
                "from cx_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and operatorcode in ('11100901','11100908','11100930','11100935','15204102','4412001028','4412001035','4412001036','4412031015','4412251010','4412255009','4412261004','4412261005','4412261007','4412845007','4412831021')\n" +
                "and substring(agentcode,5,1)<>'5'\n" +
                "and endordate >=?\n" +
                "and endordate <?\n" +
                "union\n" +
                "select '车险批单',comcode,policyno,endorseno,riskcode,inputtime,endordate,0 as sumpremium,operatorcode,'车p' as  licenseno\n" +
                "from cx_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and operatorcode in ('4412831025')\n" +
                "and substring(agentcode,5,1)<>'5'\n" +
                "and comcode = '44128314'\n" +
                "and endordate >=?\n" +
                "and endordate <?\n" +
                "union\n" +
                "select '车险批单',comcode,policyno,endorseno,riskcode,inputtime,endordate,0 as sumpremium,operatorcode,'车p' as  licenseno\n" +
                "from cx_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and operatorcode in ('4412835047')\n" +
                "and substring(agentcode,5,1)<>'5'\n" +
                "and comcode = '44128318'\n" +
                "and endordate >=?\n" +
                "and endordate <?\n" +
                "\t\t\t\t\n" +
                "union \n" +
                "select '非车批单',comcode,policyno,endorseno,riskcode,inputtime,endordate,0 as sumpremium,operatorcode,'非车p' as  licenseno\n" +
                "from fc_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and operatorcode in ('11100901','11100908','11100930','11100935','15204102','4412001028','4412001035','4412001036','4412031015','4412251010','4412255009','4412261004','4412261005','4412261007','4412845007','4412831021')\n" +
                "and substring(agentcode,5,1)<>'5'\n" +
                "and endordate >=?\n" +
                "and endordate <?\n" +
                "union \n" +
                "select '非车批单',comcode,policyno,endorseno,riskcode,inputtime,endordate,0 as sumpremium,operatorcode,'非车p' as  licenseno\n" +
                "from fc_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and operatorcode in ('4412831025')\n" +
                "and substring(agentcode,5,1)<>'5'\n" +
                "and comcode = '44128314'\n" +
                "and endordate >=?\n" +
                "and endordate <?\n" +
                "union \n" +
                "select '非车批单',comcode,policyno,endorseno,riskcode,inputtime,endordate,0 as sumpremium,operatorcode,'非车p' as  licenseno\n" +
                "from fc_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and operatorcode in ('4412831025')\n" +
                "and substring(agentcode,5,1)<>'5'\n" +
                "and comcode = '44128314'\n" +
                "and endordate >=?\n" +
                "and endordate <?;",lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month);

        //非车商保单批单
        List<Map<String, Object>> result6 = picczq.queryForList(" select '车险保单',a.comcode,a.policyno,a.ProposalNo,a.RiskCode,a.inputtime,a.StartDate,a.SumPremium,a.operatorcode,b.licenseno \n" +
                "from cx_prpcmain a,cx_prpcitem_car b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.BusinessNature<>'5'\n" +
                "and a.OperatorCode in ('15188181','44245857','44498880','84402863','A441200085','A441200090','A441200091','A441200093','A441200094','A441200095','A441200096')\n" +
                "and a.inputtime>=?\n" +
                "and a.inputtime<?\t\n" +
                "and a.RiskCode in ('DZA','DAA','DZB')\n" +
                "and a.policyno like '%P%'\n" +
                "union \n" +
                "select '非车保单',comcode,policyno,ProposalNo,RiskCode,inputtime,StartDate,SumPremium,OperatorCode,'非车险' as  licenseno\n" +
                "from fc_prpcmain \n" +
                "where underwriteflag in ('1','3')\n" +
                "and OperatorCode in ('15188181','44245857','44498880','84402863','A441200085','A441200090','A441200091','A441200093','A441200094','A441200095','A441200096')\n" +
                "and (proposalno like 'T%' or proposalno is null)\n" +
                "and BusinessNature<>'5'\n" +
                "and inputtime >=?\t\n" +
                "and inputtime <?\t\t\n" +
                "union \n" +
                "select '车险批单',comcode,policyno,EndorseNo,riskcode,inputtime,EndorDate,0 as SumPremium,OperatorCode,'车P' as  licenseno\n" +
                "from cx_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and OperatorCode in ('15188181','44245857','44498880','84402863','A441200085','A441200090','A441200091','A441200093','A441200094','A441200095','A441200096')\n" +
                "and SUBSTRING(AgentCode,5,1)<>'5'\n" +
                "and EndorDate >=?\n" +
                "and EndorDate <?\t\t\t\t\n" +
                "union \n" +
                "select '非车批单',comcode,policyno,EndorseNo,riskcode,inputtime,EndorDate,0 as SumPremium,OperatorCode,'非车P' as  licenseno\n" +
                "from fc_prpphead\n" +
                "where underwriteflag in ('1','3')\n" +
                "and OperatorCode in ('15188181','44245857','44498880','84402863','A441200085','A441200090','A441200091','A441200093','A441200094','A441200095','A441200096')\n" +
                "and SUBSTRING(AgentCode,5,1)<>'5'\n" +
                "and EndorDate >=?\n" +
                "and EndorDate <?;",lastmonth,month,lastmonth,month,lastmonth,month,lastmonth,month);


        exportToExcel.cdtjl(result1,"车商车险出单");
        exportToExcel.cdtjl(result2,"三版车险出单");
        exportToExcel.cdtjl(result3,"非车商个代车险出单");
        exportToExcel.cdtjl(result4,"车商保单批单");
        exportToExcel.cdtjl(result5,"三版保单批单");
        exportToExcel.cdtjl(result6,"非车商个代保单批单");

        return new ResultMap(1,"",true,"","");
    }
    //邮政业务到期
    public ResultMap yzdq() throws IOException {
        picczq.update("create temporary table temp1 (\n" +
                "select  a.proposalno,a.policyno,a.riskcode,a.comcode,b.insuredname,b.identifynumber,\n" +
                "        cast(null as char(120)) bbr,cast(null as char(20)) bbrsfz,\n" +
                "        a.startdate,a.enddate,a.sumamount,a.sumpremium,max(c.modename) as modename,\n" +
                "        b.phonenumber,b.mobile\n" +
                "from fc_prpcmain a,fc_prpcinsured b,fc_prpcitemkind c\n" +
                "where a.proposalno = b.proposalno\n" +
                "and a.agentcode = '44003f100013'\n" +
                "and a.policyno is not null\n" +
                "and a.sumpremium > 0\n" +
                "and a.enddate >= ? \n" +
                "and a.enddate < ?\n" +
                "and substring(b.insuredflag,1,1)=1\n" +
                "and a.proposalno = c.proposalno\n" +
                "group by c.proposalno);",nextmonth,nexttowmonth);
        picczq.update("create index temp1 on temp1(proposalno);\n");
        picczq.update("create temporary table temp2 (\n" +
                "select a.proposalno,a.insuredflag,a.insuredname,a.identifynumber\n" +
                "from fc_prpcinsured a,temp1 b\n" +
                "where a.proposalno = b.proposalno\n" +
                "and substring(a.insuredflag,2,1)=1);");
        picczq.update("create index temp2 on temp2(proposalno);");
        picczq.update("update temp1 a,temp2 b\n" +
                "        set  a.bbr = b.insuredname\n" +
                "where a.proposalno  = b.proposalno;");
        picczq.update("update temp1 a,temp2 b\n" +
                "        set  a.bbrsfz = b.identifynumber\n" +
                "where a.proposalno  = b.proposalno;");
        picczq.update("create temporary table mobile(\n" +
                "select distinct mobile from temp1\n" +
                "where mobile >= 10000000000\n" +
                "and mobile < 20000000000\n" +
                "order by mobile);");
        picczq.update("set global group_concat_max_len = 300000;");
        picczq.update("set session group_concat_max_len = 300000;");
        List<Map<String, Object>> result1 = picczq.queryForList("select * from temp1\n" +
                "order by enddate,modename;");
        List<Map<String, Object>> result2 = picczq.queryForList("select group_concat(mobile) from mobile");
        picczq.update("drop table temp1;");
        picczq.update("drop table temp2;");
        picczq.update("drop table mobile;");
        exportToExcel.yzdq(result1);
        TxtExport.creatTxtFile(nextmonth+"-"+nexttowmonth+"到期电话号码");
        TxtExport.writeTxtFile(result2);

        return new ResultMap(1,"",true,"","");

    }
    //异地车
    public ResultMap ydc(){
        picczq.update("create temporary table temp1(\n" +
                "select a.comcode,cast(null as char(120))comname,cast(null as char(120))comcodename,a.proposalno,a.policyno,a.sumpremium,a.sumtaxpremium,\n" +
                "a.agentcode, cast(null as char(120)) agentname,b.clausetype,b.carkindcode,a.startdate,a.enddate,\n" +
                "a.underwriteflag,b.licenseno,a.operatedate,a.riskcode,b.monopolycode,b.monopolyname,b.newcarflag,\n" +
                "cast(null as datetime) policynotime,cast(null as char(120)) toubaorenname,c.insuredname,c.identifynumber,\n" +
                "c.insuredaddress,cast(null as char(20)) xinxuzhuan,b.usenaturecode,cast(null as char(20)) usenaturename,\n" +
                "cast(null as char(20))province\n" +
                "from cx_prpcmain a,cx_prpcitem_car b,cx_prpcinsured c\n" +
                "where a.policyno is not null\n" +
                "and a.policyno != ''\n" +
                "and a.proposalno = b.proposalno\n" +
                "and a.proposalno = c.proposalno\n" +
                "and startdate >= ?\n" +
                "and startdate < ?\n" +
                "and substring(c.insuredflag,2,1) = 1\n" +
                "and b.licenseno not like '粤h%'\n" +
                "and b.licenseflag = 1);\n",lastmonth,month);
        picczq.update("create index proposalno on temp1(proposalno);");
        picczq.update("update temp1 set comname = '车商'\n" +
                "        where  comcode in ('44129304','44129306','44129305','44129320','44129331','44129332','44129335','44129336','44129342');");
        picczq.update("update temp1 set comname = '端州'\n" +
                "        where comcode in ('44120100','44129301','44129302','44129303','44129311','44129315','44129317','44129318','44129319',\n" +
                "              '44129321','44129329','44129330','44129338','44129343','44129324','44120313');");
        picczq.update("update temp1 set comname = '鼎湖'\n" +
                "        where comcode like '441203%' and comcode <>'44120313';");
        picczq.update("update temp1 set comname = '广宁'\n" +
                "        where comcode like '441223%';");
        picczq.update("update temp1 set comname = '怀集'\n" +
                "        where comcode like '441224%';");
        picczq.update("update temp1 set comname = '封开'\n" +
                "        where comcode like '441225%';");
        picczq.update("update temp1 set comname = '德庆'\n" +
                "        where comcode like '441226%';");
        picczq.update("update temp1 set comname = '高要'\n" +
                "        where comcode like '441283%';");
        picczq.update("update temp1 set comname = '四会'\n" +
                "        where comcode like '441284%';");
        picczq.update("update temp1 set comname = '大旺'\n" +
                "        where comcode in ('44129323','44129325');");
        picczq.update("update temp1 set comname = '电销'\n" +
                "        where comcode in ('44129400');\n");
        picczq.update("update temp1,prpdcompany set comcodename = comcname\n" +
                "where temp1.comcode = prpdcompany.comcode;");
        picczq.update("update temp1,prpdagent set temp1.agentname = prpdagent.agentname\n" +
                "where temp1.agentcode = prpdagent.agentcode;");
        picczq.update("update temp1,prptime set temp1.policynotime = prptime.operatetimeforhis\n" +
                "where temp1.policyno = prptime.certino\n" +
                "and prptime.updatetype = 'c01';");
        picczq.update("update temp1 a,cx_prpcinsured b set a.toubaorenname = b.insuredname\n" +
                "where a.proposalno = b.proposalno\n" +
                "and substring(b.insuredflag,1,1) = 1;");
        picczq.update("create temporary table temp2(\n" +
                "select proposalno,policyno,oldpolicyno\n" +
                "from cx_prpcrenewal \n" +
                "where proposalno in (select proposalno from temp1));");
        picczq.update("update temp1 set xinxuzhuan = '转保';");
        picczq.update("update temp1 set xinxuzhuan = '新保' where newcarflag = 1;");
        picczq.update("update temp1,temp2 set temp1.xinxuzhuan = '续保' \n" +
                "where temp1.proposalno = temp2.proposalno;");
        picczq.update("update temp1\n" +
                "    set usenaturename = case temp1.usenaturecode \n" +
                "        when '000' then '不区分营业非营业'\n" +
                "        when '111' then '出租、租赁'\n" +
                "        when '112' then '城市公交'\n" +
                "        when '113' then '公路客运'\n" +
                "        when '114' then '旅游客运'\n" +
                "        when '120' then '营业货车'\n" +
                "        when '121' then '营业挂车'\n" +
                "        when '180' then '运输型拖拉机'\n" +
                "        when '190' then '其它营业车辆'\n" +
                "        when '211' then '家庭自用汽车'\n" +
                "        when '212' then '非营业企业客车'\n" +
                "        when '213' then '非营业机关、事业团体客车'\n" +
                "        when '220' then '非营业货车'\n" +
                "        when '221' then '非营业挂车'\n" +
                "        when '280' then '兼用型拖拉机'\n" +
                "        when '290' then '其它非营业车辆'\n" +
                "    end\n" +
                "where temp1.usenaturecode in ('000','111','112','113','114','120','121','180','190','211','212','213','220','221','280','290');");
        picczq.update("update temp1 set province = substring(licenseno,1,1);");
        List<Map<String,Object>> result  = picczq.queryForList("select * from temp1");
        picczq.update("drop table temp1");
        picczq.update("drop table temp2");

        exportToExcel.ydc(result);

        return  new ResultMap(1,"",true,"","");
    }
    //批改
    public ResultMap pigai(){


        picczq.update("Create temporary table cp (select  a.comcode,a.riskcode,a.endorseno,a.policyno,c.proposalno,\n" +
                "b.licenseno as licenseno1,b.clausetype,b.carkindcode,\n" +
                "a.endordate,a.endortype,b.newcarflag\n" +
                "from cx_prpphead a, \n" +
                "cx_prppitem_car b,\n" +
                "cx_prpcmain c\n" +
                "where a.applyno=b.ApplyNo\n" +
                "and a.policyno=c.policyno\n" +
                "and a.endorseno like 'ED%'\n" +
                "and a.endortype like '%72%'\n" +
                "and a.endordate >= ?" +
                "and a.endordate <= ?);",lastweek,todayminone);
        picczq.update("create index proposalno on cp(proposalno);");
        List<Map<String,Object>> result  = picczq.queryForList("select a.licenseno,b.*\t\t\t \n" +
                "from cx_prpcitem_car a,cp b\n" +
                "where a.proposalno=b.proposalno;");
        picczq.update("drop table cp;");
        exportToExcel.pigai(result);
        return  new ResultMap(1,"",true,"","");

    }


}





 