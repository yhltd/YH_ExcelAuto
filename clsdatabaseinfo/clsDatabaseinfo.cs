using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace clsdatabaseinfo
{
    public class clsWangyininfo
    {
        public string Message { get; set; }
        //属性
        public string Maichangdaima { get; set; }
        public string Input_Date { get; set; }
        public int Id { get; set; }
        public string yinhang { get; set; }
        public DateTime jiaoyiriqi { get; set; }
        public double jiefangfanshenge { get; set; }//借方发生额/元(支取)       
        public double daifangfashenge { get; set; }//贷方发生额/元(收入)  
        public string zhaiyao { get; set; }
        public string jiaoyileixing { get; set; }
        public string yongtu { get; set; }
        public string duifanghuming { get; set; }
        public string beizhu { get; set; }

        //瘦身工作2017。11.10

        //public DateTime jiaoyishijian { get; set; }

        //public double yue { get; set; }
        //public string bizhong { get; set; }

        //public string duifangzhanghao { get; set; }
        //public string duifangkaihujigou { get; set; }
        //public DateTime jizhangriqi { get; set; }


        //public string zhanghumingxibiaohao { get; set; }
        //public string qiyeliushuihao { get; set; }
        //public string pingzhengzhonglei { get; set; }
        //public string pingzhenghao { get; set; }
        //public string guanlianzhanghu { get; set; }

        ////CCX3

        ////public string duifangzhanghao { get; set; }

        //public double jiefangjine { get; set; }
        //public double daifangjine { get; set; }
        ////工行-1

        ////工行-2
        //public string benfangzhanghao { get; set; }
        //public string jiedai { get; set; }

        //public string duifangdanweimingcheng { get; set; }
        //public string gexinghuaxinxi { get; set; }
        //public string duifanghanghao { get; set; }
        ////建行-1
        //public string zhanghumingcheng { get; set; }

        ////农行-1
        //public double shouxufeizonge { get; set; }
        //public string jiaoyihangming { get; set; }
        //public string duifangshengshi { get; set; }
        //public string jiaoyishuoming { get; set; }
        //public string jiaoyifuyan { get; set; }
        //public string jiaoyifangshi { get; set; }
        ////中行-1
        //public string yewuleixing { get; set; }
        //public string fukuanrenkaihuhanghao { get; set; }
        //public string fukuanrenkaihuhangming { get; set; }
        //public string fukuanrenzhanghao { get; set; }
        //public string fukuanrenmingcheng { get; set; }
        //public string shoukuanrenkaihuhanghao { get; set; }
        //public string shoukuanrenkaihuhangming { get; set; }
        //public string shoukuanrenzhanghao { get; set; }
        //public string shoukuanrenmingcheng { get; set; }
        //public string jiaoyihuobi { get; set; }
        //public double jiaoyijine { get; set; }
        //public DateTime qixiriqi { get; set; }
        //public double huilv { get; set; }
        //public string jiaoyiliushuihao { get; set; }
        //public string kehushenqinghao { get; set; }
        //public string kehuyewubianhao { get; set; }
        //public string pingzhengleixing { get; set; }

        //public string jilubiaoshihao { get; set; }
        //public string yuliuxiang1 { get; set; }
        //public string yuliuxiang2 { get; set; }
        ////交行-1
        //public string qiyeyewubianhao { get; set; }
        ////招行-1
        //public string liushuihao { get; set; }
        //public string liuchengshilihao { get; set; }
        //public string yewumingcheng { get; set; }
        //public string yewucankaohao { get; set; }
        //public string yewuzhaiyao { get; set; }
        //public string qitazhaiyao { get; set; }
        //public string shoufufangfenhangming { get; set; }
        //public string shoufufangmingcheng { get; set; }
        //public string shoufufangzhanghao { get; set; }
        //public string shoufufangkaihuhanghao { get; set; }
        //public string shoufufangkaihuhangming { get; set; }
        //public string shoufufangkaihuhangdizhi { get; set; }
        //public string muzigongsizhanghaofenhangming { get; set; }
        //public string muzigongsizhanghao { get; set; }
        //public string muzigongsimingcheng { get; set; }
        //public string xinxibiaozi { get; set; }
        //public string youfoufujianxinxi { get; set; }
        //public string chongzhangbiaozhi { get; set; }
        //public string kuozhanzhaiyao { get; set; }
        //public string jiaoyifenxima { get; set; }
        //public string piaojuhao { get; set; }
        //public string shangwuzhifudingdanhao { get; set; }
        //public string neibubianhao { get; set; }
        //public string jiaoyishijian1 { get; set; }
        ////兴业
        //public string yinhangliushuihao { get; set; }
        //public string zhanghao { get; set; }
        //public string huming { get; set; }
        //public string xianzhuan { get; set; }
        //public string duifangyinhang { get; set; }
        ////中信
        //public long zhujijiaoyima { get; set; }
        //public string guiyuanjiaoyima { get; set; }
        //public string beichongzhangbiaozhi { get; set; }
        //public string feijinrongbiaozhi { get; set; }
        //public string zhaiyaodaima1 { get; set; }
        //public string zhaiyaodaima2 { get; set; }
        //public string zhidanyuanid { get; set; }
        //public string zhidanchaozuoyuanxingming { get; set; }
        //public string fuheID { get; set; }
        //public string fuheyuanxingming { get; set; }
        //public string waihangzhanghumingcheng { get; set; }
        //public string waihangkaihuhangmingcheng { get; set; }
        //public string tuipiaobiaoshi { get; set; }
        //public DateTime tuipiaoriqi { get; set; }
        //public string tuipiaochangci { get; set; }
        ////j锦州
        //public string jiaoyimiaosu { get; set; }

        ////杭州
        //public string xuhao { get; set; }
        //public string fuyan { get; set; }
        ////重庆农商银行
        //public string huiruhuichu { get; set; }
        //public string qudao { get; set; }
        ////

        ////
        public string zhifufangshi_leibie { get; set; }
    }
    public class clsribaodatasoureinfo
    {
        public string Message { get; set; }
        public DateTime riqi { get; set; }
        public double diaopaijia_zhengjia { get; set; }
        public double POSjine { get; set; }
        public double shuaka { get; set; }
        public double xianjin { get; set; }
        public double zekoujine { get; set; }
        public double jingwaika { get; set; }
        public double jiansheyinhang { get; set; }
        public double xianjine { get; set; }
        public double lianhuaokka { get; set; }
        public double simateka { get; set; }
        public double okka { get; set; }
        public double yikatong { get; set; }
        public double wuxianka { get; set; }
        public double weikangka { get; set; }
        public double shangchangka { get; set; }
        public double weixin { get; set; }
        public double zhifubao { get; set; }
        public double zhanghujine { get; set; }
        public double posyushijishouruchayiyuanyin { get; set; }
        public string beizhu { get; set; }
        public double zengpinjianshu { get; set; }
        public double zidaijianshu { get; set; }
        public double jifen { get; set; }
        //属性
        public string Maichangdaima { get; set; }
        public string QISHU { get; set; }
        public string mingcheng { get; set; }
        public double leijichayishu { get; set; }
        public double shangyuechayi { get; set; }
        public string j4 { get; set; }
        public string k4 { get; set; }
        public string l4 { get; set; }
        public string m4 { get; set; }
        public string diqudaima { get; set; }
        public string Input_Date { get; set; }
        public int Id { get; set; }
        // 手续费

        public double NCshouxufei { get; set; }
        public double yingfankuan { get; set; }
        public double shijifankuan { get; set; }
        public double chae { get; set; }
        //现金
        public double chae_xianjin { get; set; }
        public double shijicunxian { get; set; }
        //代收银-1
        public double zhangkoufeiyong { get; set; }
        //手续费，现金保存用到
        public string leixing { get; set; }
        //实际返款-账扣
        public double shijifankuanzhangkou { get; set; }

        //
        public double zhanghujinesum { get; set; }
        //记录手续费的支付方式 构造table

        public string zhifufangshi_leibie { get; set; }
    }
    public class clsKEYinfo
    {
        public string Message { get; set; }
        public string diqu { get; set; }
        public string fuzeren { get; set; }

        public string maichangname { get; set; }
        public string Maichangdaima { get; set; }
        public string naishuidanwei { get; set; }

        public string yinhangjiancheng { get; set; }
        public string yinhangkemu { get; set; }
        public string shoukuanfangshi { get; set; }

        public string duiyinglie { get; set; }
        public string keytext { get; set; }
        public string Input_Date { get; set; }
    }
    public class clsbanktelkeyinfo
    {
        public string Message { get; set; }
        public string Demo { get; set; }
        public string name { get; set; }
        public string cloumnindex { get; set; }
    }
    public class clsHandingChargeinfo
    {
        public string Message { get; set; }
        public string faren { get; set; }
        public string maichangname { get; set; }

        public string nashuidanwei { get; set; }
        public string Maichangdaima { get; set; }
        public string shuakashouxufei { get; set; }

        public string weixinshouxufei { get; set; }
        public string zhifubaoshouxufei { get; set; }
        public string shangchangkashouxufei { get; set; }

        public string okkashouxufei { get; set; }
        public string jingwaikashouxufei { get; set; }
        public string siweikashouxufei { get; set; }
        public string Input_Date { get; set; }
    }
    public class clsUSERinfo
    {
        public string Message { get; set; }
        public string userid { get; set; }
        public string passwordid { get; set; }
        public DateTime Input_Date { get; set; }
        public string Readid { get; set; }
        public string Writeid { get; set; }
        public string Adminid { get; set; }
        public string salse_code { get; set; }
        public int Id { get; set; }
    }
    public class clsR2Rbankchargeinfo
    {
        public string Message { get; set; }
        public string yinhang { get; set; }
        public string lieming { get; set; }
        public string Input_Date { get; set; }
        public string miaoshu1 { get; set; }
        public string miaoshu2 { get; set; }
        public string miaoshu3 { get; set; }
        //
        public string type_name { get; set; }
        public int Id { get; set; }
    }
    public class clsR2RBankinfo
    {
        public string Message { get; set; }
        public string yinhang { get; set; }
        public DateTime jiaoyiriqi { get; set; }
        public double jiefang { get; set; }
        public double daifang { get; set; }
        public string zhaiyao { get; set; }
        public string yongtu { get; set; }


        public string Input_Date { get; set; }
        public int Id { get; set; }
        //new 
        public string nashuidanwei { get; set; }
        public string xuhao { get; set; }
        public string gongsidaima { get; set; }
        public string type_name { get; set; }//区分手续费和利息

    }
    public class clsR2RBankMappinginfo
    {
        public string Message { get; set; }
        public string yinhang { get; set; }
        public string jiaoyiriqi { get; set; }
        public string jiefang { get; set; }
        public string daifang { get; set; }
        public string zhaiyao { get; set; }
        public string yongtu { get; set; }

        public int Id { get; set; }
        public string Input_Date { get; set; }

    }
    public class clsR2Rtax_unitinfo
    {
        public string Message { get; set; }
        public string jizhangma { get; set; }
        public string kemu { get; set; }
        public string nashuidanwei { get; set; }
        public string chengbenzhongxin { get; set; }
        public string yinglizhongxin { get; set; }
        public string zijinzhongxin { get; set; }
        public string faren { get; set; }

        public int Id { get; set; }
        public string Input_Date { get; set; }


    }

    public class clsAccrual_KAEVinfo
    {
        public string Message { get; set; }
        public string ka { get; set; }
        public string ev { get; set; }


        public int Id { get; set; }
        public string Input_Date { get; set; }

    }


    public class clsR2RTaxpaid_keyinfo
    {
        public string Message { get; set; }
        public string nashuidanwei { get; set; }
        public string lirunzhongxin { get; set; }

        public string Maichangdaima { get; set; }
        public double yingjiaozhengzhishuijine { get; set; }
        public string yingjiaozhengzhishui_kemu { get; set; }
        public double yingjiaofujiashui { get; set; }//金额
        public string yingjiaofujiashui_kemu { get; set; }
        public string other { get; set; }
        public string other_kemu { get; set; }
        public double other_jine { get; set; }
        public int Id { get; set; }
        //
        public string lie_name { get; set; }
        public string Input_Date { get; set; }
    }

    public class clsR2Rshujiininfo
    {
        public string Message { get; set; }
        public string xuhao { get; set; }
        public string xiang { get; set; }
        public string jizhangma { get; set; }
        public string kemu { get; set; }
        public string nashuidanwei { get; set; }
        public string chengbenzhongxin { get; set; }
        public string yinglizhongxin { get; set; }
        public string zijinzhongxin { get; set; }
        public string faren { get; set; }
        public string gaiyao { get; set; }
        public double jiaoyihuobijine { get; set; }
        public int Id { get; set; }
        public string Input_Date { get; set; }
        //
        public string jiti_nashuitype { get; set; }

    }
    public class clsR2accrualsapinfo
    {
        public string Message { get; set; }
        public int Id { get; set; }
        public string Input_Date { get; set; }
        public string yiqingxiangmu { get; set; }
        public string jizhangriqi { get; set; }
        public string pingzhengriqi { get; set; }
        public string nashuidanwei { get; set; }
        public string lirunzhongxin { get; set; }
        public string benbijine { get; set; }
        public string wenben { get; set; }
        public string pingzhengleixing { get; set; }
        public string jizhangdaima { get; set; }
        public string pingzhenghaoma { get; set; }
        public string benbi { get; set; }
        public string qingzhangpingzheng { get; set; }
        public string glkemu { get; set; }
        public string nianduyuefen { get; set; }
        public string yonghumingcheng { get; set; }
        public string congxiao { get; set; }
        //
        public string leiming { get; set; }

        public string xuhao { get; set; }
        public string gaiyao { get; set; }
        //gongsidaima
        public string gongsidaima { get; set; }
        //
        //faren
        public string faren { get; set; }
        //选择要分析的日期
        public string selecttime { get; set; }
        // 标记已计算过的KA 
        public string is_mergeKA { get; set; }
    }
    public class clsStatus_Clickinfo
    {
        public string Message { get; set; }
        public string ClickType { get; set; }
        public string Comments { get; set; }
        public string whose { get; set; }
        public DateTime Input_Date { get; set; }
        public int Id { get; set; }
    }
    public class clsToleranceinfo
    {
        public string Message { get; set; }
        public string faren { get; set; }
        public string gudingzujin_zheng { get; set; }
        public string gudingzujin_fu { get; set; }
        public string ticengzujin_zheng { get; set; }
        public string ticengzujin_fu { get; set; }
        public string poszujin_zheng { get; set; }
        public string poszujin_fu { get; set; }
        public string wuyefei_zheng { get; set; }
        public string wuyefei_fu { get; set; }
        public string tuiguangfei_zheng { get; set; }
        public string tuiguangfei_fu { get; set; }
        public string zujin_zheng { get; set; }
        public string zujin_fu { get; set; }


        public string Input_Date { get; set; }
        public int Id { get; set; }

    }
    public class clsYutixinxihuizonginfo
    {
        public string Message { get; set; }
        public string Maichangdaima { get; set; }
        public string mingcheng { get; set; }
        public string lirunzhongxin { get; set; }
        public string gudingzujin { get; set; }
        public string shoudaofapiao { get; set; }
        public string ticengzujin { get; set; }
        public string poszujin { get; set; }
        public string wuyefei { get; set; }
        public string shuidianfei { get; set; }
        public string tuiguangfei { get; set; }

        public string POSshoudaofapiao { get; set; }
        public string Wuyefei_shoudaofapiao { get; set; }
        public string tuiguang_shoudaofapiao { get; set; }


    }
    public class clsAP_fuwuinfo
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string CaseID { get; set; }
        public string biaoti { get; set; }
        public string shenqingshijian { get; set; }
        public string shenqingzhe { get; set; }
        public string zhuangtai { get; set; }
        public string chulizhe { get; set; }
        public string fuwufenleida { get; set; }
        public string fuwufenleizhong { get; set; }
        public string fuwufenleixiao { get; set; }
        public string tiaoxingma { get; set; }
        public string youxianshunxu { get; set; }
        public string gongsidaima { get; set; }
        public string pongzhengzubianhao { get; set; }
        public string sapshenpizhuangtai { get; set; }
        public string chulituandui { get; set; }
        public string wanchengshijian { get; set; }
        public string fujian { get; set; }
        public string dianshishenpizhuangtai { get; set; }
        public string dianzishenpishijian { get; set; }
        public string pinpai { get; set; }
        public string maichangdaima { get; set; }
        public string maichangming { get; set; }
        public string kaiyeriqi { get; set; }
        public string maichangdengji { get; set; }
        public string zuijinzhuanghuangriqi { get; set; }
        public string shenqingfarendaima { get; set; }
        public string manyidupingjia { get; set; }
        public string other_29 { get; set; }
        public string OWER { get; set; }
        public string Input_Date { get; set; }
        public DateTime Single_Date { get; set; }
    }
    public class clsAP_Single_info
    {
        public string Message { get; set; }
        public string CaseID { get; set; }
        public string shenqingzhe { get; set; }
        public string fapiaoleixing { get; set; }
        public string gongsidaima { get; set; }
        public string pingzhengzubianhao { get; set; }
        public string OWER { get; set; }

        //new
        public string aType { get; set; }
        public int Volume { get; set; }
        public string changelogTO { get; set; }
        public string Input_Date { get; set; }

    }
    public class clsAPlegalinfo
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string legal_person { get; set; }
        public string legal_type { get; set; }
        public string Input_Date { get; set; }
    }
    public class clsPRC_zicanfuzaiinfo
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string D38 { get; set; }
        public string H38 { get; set; }
        public string FileName { get; set; }
        public string H35 { get; set; }
        public string G35 { get; set; }
        public string D28 { get; set; }
    }
    public class clsTAX_carryinfo
    {
        public int Id { get; set; }
        public string Message { get; set; }
        //增值税结转科目 sheet 
        public string shujinkemu { get; set; }

        //验证-1 sheet
        public string naishuidanwei { get; set; }

        public string liruzhongxin { get; set; }

        public string xianshihuobijine { get; set; }


    }
    public class clsZFIR5080info
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string pingzhenghaoma { get; set; }
        public string fapiaohao { get; set; }
        public string benbijine { get; set; }

    }
    public class cls13450101info
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string pingzhenghaoma { get; set; }
        public string fapiaohao { get; set; }
        public string benbijine { get; set; }
        public string nashuidanwei { get; set; }
        public string pingzhengleixing { get; set; }
        public string linrunzhongxin { get; set; }

        public string chongxiao { get; set; }
        public string gongsidaima { get; set; }
     

    }
    public class clsRenzhengmingxi
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string xuhao { get; set; }
        public string fapiaodaima { get; set; }
        public string fapiaohao { get; set; }
        public string kaipiaoriqi { get; set; }
      

        public string xiaofangmingcheng { get; set; }
        public string jine { get; set; }
       public string shuie { get; set; }
        public string fapiaozhuangtai { get; set; }
        public string fapiaoleixing { get; set; }
        public string naishuishibiehao { get; set; }
        public string shuikuaisuoshuqi { get; set; }


        /// <summary>
        /// 附加
        /// </summary>
        public string sheetname { get; set; }
        public string nashuidanwei { get; set; }
        public string Input_Date { get; set; }
   
    }
    public class clsjinxiangkemutiaozhengsheet_info
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string gongsidaima { get; set; }
        public string jizhangdaima { get; set; }
        public string gl_kemu { get; set; }
        public string jizhangriqi { get; set; }
        public string pingzhengriqi { get; set; }
        public string pingzhengleixing { get; set; }
        public string pingzhenghaoma { get; set; }
        public string nashuidanwei { get; set; }
        public string lirunzhongxin { get; set; }
        public string benbijine { get; set; }
        public string wenben { get; set; }
        public string congxiao { get; set; }
        public string yonghumingcheng { get; set; }
        public string Input_Date { get; set; }
   

    }
    public class clsXiaofangmingcheng_info
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string name { get; set; }
        public string Input_Date { get; set; }
    }
    public class clsRenzhengmingxinLine_info
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string faren { get; set; }
        public string nashuidanwei { get; set; }
        public string chengbenzhongxin { get; set; }
        public string yingli_zijinzhongxin { get; set; }
        public string Input_Date { get; set; }
    }
    public class clsZFIR6030_info
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string mugongsiA { get; set; }
        public string gukeB { get; set; }
        public string kehumingC { get; set; }
        public string xiaoshouzaiquanD { get; set; }
        public string nashuibiaozhunE { get; set; }
        public string shuieF { get; set; }
        public string caigouozaiwuG{ get; set; }
        public string caigouozaiwuH { get; set; }
        public string zigongsiI { get; set; }
        public string gongyingshangJ { get; set; }
        public string gongyingshangmingK { get; set; }
        public string caigouozaiwuL { get; set; }
        public string naishuibiaozhunM { get; set; }
        public string shuiN { get; set; }
        public string caigouozaiwuO { get; set; }
        public string caigouozaiwuP { get; set; }
        public string xiaoshoucaigouchaQ { get; set; }
        public string keshuibiaozhunchaR { get; set; }
        public string zhengzhishuichaS { get; set; }
        public string zaiquanheshuanpingzhengT { get; set; }
        public string zaiquanheshuanpingzhengU { get; set; }
        public string zaiquanheshuanpingzhengV { get; set; }
        public string zaiquanheshuanpingzhengW { get; set; }
        public string xiaoxiwenbenX { get; set; }
        //
        public string  nashuidanwei { get; set; }
       
    }
    public class clsLine_info
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public string kemu { get; set; }
        public string jizhangma { get; set; }
        public string nashuidanwei { get; set; }
        public string jiaoyihuobishuie { get; set; }
        public string jiaoyihuobijine { get; set; }
    }
}
