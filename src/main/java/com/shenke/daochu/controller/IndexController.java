package com.shenke.daochu.controller;

import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.shenke.daochu.util.DaoUtil;
import com.shenke.daochu.util.LogUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@ResponseBody
@Controller
public class IndexController {
    private String path = "d:/Test.xlsx";

    @RequestMapping("/findAll")
    public void findAll() {
        String sql = "select \n" +
                "mzh as 门诊号,\n" +
                "sfzy as 是否住院,\n" +
                "zyh as 住院号,\n" +
                "fz as 是否复诊,\n" +
                "xingming as 患者姓名,\n" +
                "xingbie as 患者性别,\n" +
                "jianhurenxm as 监护人姓名,\n" +
                "chushenny as 出生年月,\n" +
                "nianling as 年龄,\n" +
                "lianxifs as 联系方式,\n" +
                "bingrensy as 患者属于,\n" +
                "xianzhuzis as 现住址省,\n" +
                "xianzhuzishi as 现住址市区,\n" +
                "xianzhuzix as 现住址乡镇,\n" +
                "xianzhuzhixx as 现住址详细,\n" +
                "zhiyesjet as zhiyesjet,\n" +
                "zhiyetyet as zhiyetyet,\n" +
                "zhiyexs as zhiyexs,\n" +
                "zhiyenm as zhiyenm,\n" +
                "zhiyemg as zhiyemg,\n" +
                "zhiyesyfw as zhiyesyfw,\n" +
                "zhiyecyspy as zhiyecyspy,\n" +
                "zhiyegr as zhiyegr,\n" +
                "zhiyeywry as zhiyeywry,\n" +
                "zhiyegbzy as zhiyegbzy,\n" +
                "zhiyeltry as zhiyeltry,\n" +
                "zhiyejs as zhiyejs,\n" +
                "zhiyejwjdy as zhiyejwjdy,\n" +
                "zhiyemm as zhiyemm,\n" +
                "zhiyeym as zhiyeym,\n" +
                "zhiyeqt as zhiyeqt,\n" +
                "zhiyeqtxx as zhiyeqtxx,\n" +
                "zhiyebx as zhiyebx,\n" +
                "fabingsj as 发病时间,\n" +
                "jiuzhensj as 就诊时间,\n" +
                "qsfs as 发烧,\n" +
                "qsfsssd as 温度,\n" +
                "qsmsch as 面色潮红,\n" +
                "qsmscb as 面色苍白,\n" +
                "qsfg as 发绀,\n" +
                "qsts as 脱水,\n" +
                "qskk as 口渴,\n" +
                "qsfz as 浮肿,\n" +
                "qstzxj as 体重下降,\n" +
                "qshz as 寒战,\n" +
                "qsfl as 乏力,\n" +
                "qspx as 贫血,\n" +
                "qszz as 肿胀,\n" +
                "qssm as 失眠,\n" +
                "qswg as 畏光,\n" +
                "qskyhw as 口有糊味,\n" +
                "qsjsw as 金属味,\n" +
                "qsfzxw as 肥皂咸味,\n" +
                "qstygd as 唾液过多,\n" +
                "qszwxc as 足腕下垂,\n" +
                "qscz as 色素沉着,\n" +
                "qstp as 脱皮,\n" +
                "qszjcxbd as 指甲出现白带,\n" +
                "qsqt as 全身其他,\n" +
                "qsqtxx as 全身其他详情,\n" +
                "xhex as 恶心,\n" +
                "xhot as 呕吐,\n" +
                "xhoucs as 呕吐次数,\n" +
                "xhft as 腹痛,\n" +
                "xhbm as 便秘,\n" +
                "xhljhz as 里急后重,\n" +
                "xhqt as 消化其他,\n" +
                "xhqtxx as 消化其他详情,\n" +
                "xhfx as 腹泻,\n" +
                "xhfxcs as 腹泻次数,\n" +
                "xhxb as xhxb,\n" +
                "xhsyb as xhsyb,\n" +
                "xhmgyb as xhmgyb,\n" +
                "xhnyb as xhnyb,\n" +
                "xhnxb as xhnxb,\n" +
                "xhxryb as xhxryb,\n" +
                "xhxxyb as xhxxyb,\n" +
                "xhhb as xhhb,\n" +
                "xhdbqt as xhdbqt,\n" +
                "xhdbqtxx as xhdbqtxx,\n" +
                "hxdc as 呼吸短促,\n" +
                "hxgx as 咯血,\n" +
                "hxkn as 呼吸困难,\n" +
                "hxqt as 呼吸其他,\n" +
                "hxqtxx as 呼吸其他详情,\n" +
                "xnxm as 胸闷,\n" +
                "xnxt as 胸痛,\n" +
                "xnxj as 心悸,\n" +
                "xnqd as 气短,\n" +
                "xnqt as 心脑血管其他,\n" +
                "xnqtxx as 心脑血管其他详情,\n" +
                "mnnljs as 尿量减少,\n" +
                "mnbbsqtt as 背部肾区疼痛,\n" +
                "mnsjs as 肾结石,\n" +
                "mnnzdx as 尿中带血,\n" +
                "mnqt as 泌尿其他,\n" +
                "mnqtxx as 泌尿其他详情,\n" +
                "sjtt as 头痛,\n" +
                "sjhm as 昏迷,\n" +
                "sjjj as 惊厥,\n" +
                "sjzw as 谵妄,\n" +
                "sjth as 瘫痪,\n" +
                "sjyykn as 言语困难,\n" +
                "sjtykn as 吞咽困难,\n" +
                "sjgjyc as 感觉异常,\n" +
                "sjjssc as 精神失常,\n" +
                "sjfs as 复视,\n" +
                "sjslmh as 视力模糊,\n" +
                "sjxy as 眩晕,\n" +
                "sjyjxc as 眼睑下垂,\n" +
                "sjztmm as 肢体麻木,\n" +
                "sjmsgjza as 末梢感觉障碍,\n" +
                "sjtkyc as 瞳孔异常,\n" +
                "sjtkkd as 瞳孔扩大,\n" +
                "sjtkgd as 瞳孔固定,\n" +
                "sjtkss as 瞳孔收缩,\n" +
                "sjzcg as 针刺感,\n" +
                "sjcc as 抽搐,\n" +
                "sjqt as 神经其他,\n" +
                "sjqtxx as 神经其他详情,\n" +
                "pfsy as 瘙痒,\n" +
                "pfszg as 烧灼感,\n" +
                "pfpz as 皮疹,\n" +
                "pfcxd as 出血点,\n" +
                "pfhd as 黄疸,\n" +
                "pfqt as 皮肤其他,\n" +
                "pfqtxx as 皮肤其他详情,\n" +
                "jxcwy as 急性胃肠炎,\n" +
                "grxfx as 感染性腹泻,\n" +
                "dmgzd as 毒蘑菇中毒,\n" +
                "cdzd as 菜豆中毒,\n" +
                "hdzd as 河豚中毒,\n" +
                "rdzd as 肉毒中毒,\n" +
                "yxdyzd as 亚硝酸盐中毒,\n" +
                "nyzd as nyzd,\n" +
                "cbzdqt as 初步诊断其他,\n" +
                "cbzdqtxx as 初步诊断其他详情,\n" +
                "sfsykss as 是否使用抗生素,\n" +
                "kssmc as kssmc,\n" +
                "jwsw as 无既往病史,\n" +
                "jwsybxhdyz as jwsybxhdyz,\n" +
                "jwskleb as 克罗恩病,\n" +
                "jwsxhdky as 消化道溃疡,\n" +
                "jwsxhdzl as 消化道肿瘤,\n" +
                "jwscyjzhz as 肠易激综合征,\n" +
                "jwsnmy as 脑膜炎,\n" +
                "jwsnzl as 脑肿瘤,\n" +
                "jwsqt as 既往病史其他,\n" +
                "jwsqtxx as 既往病史其他详情,\n" +
                "blxx as 暴露信息,\n" +
                "spmc1 as spmc1,\n" +
                "spfl1 as spfl1,\n" +
                "jgfs1 as jgfs1,\n" +
                "sppp1 as sppp1,\n" +
                "sccj1 as sccj1,\n" +
                "jscs1 as jscs1,\n" +
                "jslx1 as jslx1,\n" +
                "gmdd1 as gmdd1,\n" +
                "jsy1 as jsy1,\n" +
                "jsr1 as jsr1,\n" +
                "jss1 as jss1,\n" +
                "jsrs1 as jsrs1,\n" +
                "qtrsffb as qtrsffb,\n" +
                "spmc2 as spmc2,\n" +
                "spfl2 as spfl2,\n" +
                "jgfs2 as jgfs2,\n" +
                "sppp2 as sppp2,\n" +
                "sccj2 as sccj2,\n" +
                "jscs2 as jscs2,\n" +
                "jslx2 as jslx2,\n" +
                "gmdd2 as gmdd2,\n" +
                "jsy2 as jsy2,\n" +
                "jsr2 as jsr2,\n" +
                "jss2 as jss2,\n" +
                "jsrs2 as jsrs2,\n" +
                "qtrsffb2 as qtrsffb2,\n" +
                "spmc3 as spmc3,\n" +
                "spfl3 as spfl3,\n" +
                "jgfs3 as jgfs3,\n" +
                "sppp3 as sppp3,\n" +
                "sccj3 as sccj3,\n" +
                "jscs3 as jscs3,\n" +
                "jslx3 as jslx3,\n" +
                "gmdd3 as gmdd3,\n" +
                "jsy3 as jsy3,\n" +
                "jsr3 as jsr3,\n" +
                "jss3 as jss3,\n" +
                "jsrs3 as jsrs3,\n" +
                "qtrsffb3 as qtrsffb3,\n" +
                "keshibm as keshibm\n" +
                "from fyk_main";
        System.out.println("查询");
        Map<String, Object> map = new HashMap<>();
        Connection connection = DaoUtil.getConnection();
        try {
            PreparedStatement preparedStatement = connection.prepareStatement(sql);
            ResultSet resultSet = preparedStatement.executeQuery();
            List<Map<String, Object>> mapList = DaoUtil.getresultSet(resultSet);
            ExcelWriter writer = ExcelUtil.getWriter(path);
            writer.write(mapList, true);
            writer.flush();
            writer.close();
        } catch (Exception e) {
            LogUtil.printLog(e, e.getClass());
        }
    }

    /***
     * 文件下载
     * @param request
     * @param response
     * @return
     */
    @RequestMapping("/download")
    public void downloadFile(HttpServletRequest request, HttpServletResponse response) {
        System.out.println("导出");
        String fileName = "Test.xlsx";
        if (fileName != null) {
            //当前是从该工程的WEB-INF//File//下获取文件(该目录可以在下面一行代码配置)然后下载到C:\\users\\downloads即本机的默认下载的目录
            String realPath = "d:/";
            File file = new File(realPath, fileName);
            if (file.exists()) {
                response.setContentType("application/force-download");// 设置强制下载不打开
                response.addHeader("Content-Disposition",
                        "attachment;fileName=" + fileName);// 设置文件名
                byte[] buffer = new byte[1024];
                FileInputStream fis = null;
                BufferedInputStream bis = null;
                try {
                    fis = new FileInputStream(file);
                    bis = new BufferedInputStream(fis);
                    OutputStream os = response.getOutputStream();
                    int i = bis.read(buffer);
                    while (i != -1) {
                        os.write(buffer, 0, i);
                        i = bis.read(buffer);
                    }
                    LogUtil.printLog("下载成功");
                } catch (Exception e) {
                    e.printStackTrace();
                    LogUtil.printLog(e, e.getClass());
                } finally {
                    if (bis != null && fis != null) {
                        try {
                            bis.close();
                            fis.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                            LogUtil.printLog(e, e.getClass());
                        }
                    }
                }
            }
        }
    }
}