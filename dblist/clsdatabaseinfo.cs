using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

namespace dblist
{
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

        public string keytext11 { get; set; }
        public string keytext12 { get; set; }
    }
    public class clsTitleinfo
    {
        public string Message { get; set; }
        public string huozhubiaohaoA { get; set; }
        public string huozhumingchengB { get; set; }

        public string xiangmubianmaC { get; set; }
        public string xiangmumingchengD { get; set; }
        public string danjiaE { get; set; }

        public string chanshengriqiF { get; set; }
        public string jifeiliangG { get; set; }
        public string jineH { get; set; }

        public string youhuijineI    { get; set; }
        public string kaishiriqiJ { get; set; }

        public string jieshuriqiK { get; set; }
      
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

    public class clsPDFInfo
    {
        public string Message { get; set; }
        public string PDF_File_Name { get; set; }
        public string PDF_File_pag { get; set; }

        #region Excel
        public string LeaseScheduleNo { get; set; }
        public string BillingPeriod { get; set; }
        public string InvNo { get; set; }
        public string InvAmtexgst { get; set; }
        public string InvAmtincgst { get; set; }
        public string SmartbuyNo { get; set; }
        #endregion
        public string Description { get; set; }
        public string TotalAmount { get; set; }
        //PDF
        public string PDFinfo { get; set; }
        public string Customername { get; set; }
        public string xintuogongsijiancheng { get; set; }
        public string xintuogongsiquancheng { get; set; }
        public string chengligongsi { get; set; }
        public string zhucedishengfen { get; set; }
        public string zhucedizhi { get; set; }
        public string suoshuyinjianju { get; set; }
        public string xinyeyinhanggongfenyouxiangongsi { get; set; }
        public string aodaliyaguomingyinhang { get; set; }
        public string fujianshengnengyuanjituanyouxiangongsi { get; set; }
        public string dongshizhang { get; set; }
        public string zongjinli { get; set; }
        public string huobizijinA9 { get; set; }
        public string huokuanA10 { get; set; }
        public string jiaoyixingjinrongzichanA11 { get; set; }
        public string kegongchushoujinrongzichanA12 { get; set; }
        public string chiyouzhidaoqitouziA13 { get; set; }
        public string changqiguquantouziA14 { get; set; }
        public string qitaA15 { get; set; }
        public string hejiA16 { get; set; }
        public string jichuchanyeA18 { get; set; }
        public string fangdichanA19 { get; set; }
        public string zhengjuanshichangA20 { get; set; }
        public string gongshangqiyeA21 { get; set; }
        public string jinrongjigouA22 { get; set; }
        public string qitaA23 { get; set; }
        public string hejiA24 { get; set; }
        public string jiheA27 { get; set; }
        public string jiheA27_F27 { get; set; }
        public string danyiA28 { get; set; }
        public string danyiA28_F27 { get; set; }
        public string caichanquanA29 { get; set; }
        public string hejiA30 { get; set; }
        public string zhengjuanleiA33 { get; set; }
        public string zhengjuanleiA33_F41 { get; set; }

        public string guquanleiA34 { get; set; }
        public string guquanleiA34_F34 { get; set; }
        public string rongzileiA35 { get; set; }
        public string rongzileiA35_F43 { get; set; }



        public string qitaleiA36 { get; set; }
        public string shiwuguanlileiA37 { get; set; }
        public string shiwuguanlileiA37_F37 { get; set; }



        public string hejiA38 { get; set; }


        public string zhengjuanleiA41 { get; set; }
        public string zhengjuanleiA41_F41 { get; set; }



        public string guquanleiA42 { get; set; }
        public string guquanleiA42_F41 { get; set; }

        public string rongzileiA43 { get; set; }
        public string rongzileiA43_F41 { get; set; }


        public string qitaleiA44 { get; set; }
        public string shiwuguanlileiA45 { get; set; }

 
        public string hejiA46 { get; set; }
        public string jiheA49 { get; set; }
        public string jiheA49_G49 { get; set; }
        public string jiheA49_H49 { get; set; }



        public string danyiA50 { get; set; }
        public string danyiA50_G50 { get; set; }
        public string danyiA50_H50 { get; set; }



        public string caichanquanA51 { get; set; }
        public string caichanquanA51_F51 { get; set; }
        public string caichanquanA51_G51 { get; set; }




        public string hejiA52 { get; set; }
        public string zhengjuanleizhudongA53 { get; set; }
        public string zhengjuanleizhudongA53_G53 { get; set; }
        public string zhengjuanleizhudongA53_H53 { get; set; }


        public string guquanleizhudongA54 { get; set; }
        public string guquanleizhudongA54_G54 { get; set; }
        public string guquanleizhudongA54_H54 { get; set; }


        public string rongzileizhudongA55 { get; set; }
        public string rongzileizhudongA55_G55 { get; set; }
        public string rongzileizhudongA55_H55 { get; set; }



        public string qitaleizhudongA56 { get; set; }
        public string shiwuguanlileizhudongA57 { get; set; }
        public string shiwuguanlileizhudongA57_G57 { get; set; }
        public string shiwuguanlileizhudongA57_H57 { get; set; }




        public string zhudonghejiA58 { get; set; }
        public string zhengjuanleibeidongA59 { get; set; }
        public string zhengjuanleibeidongA59_F59 { get; set; }
        public string zhengjuanleibeidongA59_H59 { get; set; }



        public string guquanleibeidongA60 { get; set; }
        public string guquanleibeidongA60_F60 { get; set; }
        public string guquanleibeidongA60_H60 { get; set; }

        public string rongzileibeidongA61 { get; set; }
        public string rongzileibeidongA61_F61 { get; set; }
        public string rongzileibeidongA61_H61 { get; set; }


        public string qitaleibeidongA62 { get; set; }
        public string shiwuguanlileibeidongA63 { get; set; }
        public string shiwuguanlileibeidongA63_F63 { get; set; }
        public string shiwuguanlileibeidongA63_H63 { get; set; }



        public string beidonghejiA64 { get; set; }
        public string jieheA67 { get; set; }
        public string jieheA67_F67 { get; set; }
        public string danyiA68 { get; set; }
        public string danyiA68_F68 { get; set; }

        public string caichanquanA69 { get; set; }
        public string xinzhenghejiA70 { get; set; }
        public string zhudongguanliA71 { get; set; }
        public string zhudongguanliA71_F71 { get; set; }
        public string beidongguanliA72 { get; set; }
        public string shouxufeijiyongjinA73 { get; set; }
        public string shouxufeijiyongjinA75 { get; set; }
        public string lixishouruA76 { get; set; }
        public string touzishouyiA77 { get; set; }
        public string qizhongguquanA78 { get; set; }
        public string qizhongzhengjuanA79 { get; set; }
        public string qizhongqitaA80 { get; set; }
        public string gongyunjiazhibiandongshunyiA81 { get; set; }
        public string qitaA82 { get; set; }

        public string shouruhejiA83 { get; set; }
        public string shujinfujianA84 { get; set; }
        public string yewujiguanlifeiA85 { get; set; }
        public string zichanjianzhishunshiA86 { get; set; }
        public string qitajingchengbenhuoshunshiA87 { get; set; }
        public string lirunzongeA88 { get; set; }
        public string jinglirunA89 { get; set; }

        //右侧 单位：亿元 	 2018年 	 2017年 	 2016年 	 2015年 	 2014年 

        public string zhucezibenJ2 { get; set; }
        public string guyouzichanJ3 { get; set; }
        public string xintuozichanJ4 { get; set; }
        public string qizhongjiheJ5 { get; set; }
        public string qizhongzhudongJ6 { get; set; }
        public string qizhongzhudongrongzileiJ7 { get; set; }
        public string xinzhengguimoJ8 { get; set; }
        public string xinzhengjiheJ9 { get; set; }
        public string xinzengzhudongJ10 { get; set; }
        public string qingsuanguimoJ11 { get; set; }
        public string shouruJ12 { get; set; }
        public string jinglirunJ13 { get; set; }
        public string renjunshouruJ14 { get; set; }
        public string renjunyewugaunlifeiJ15 { get; set; }
        public string J16renjunjingliru { get; set; }
        // 职工人数 	 2018年 	 占比 	 2017年 	 占比 	 2016年 

        public string J20guanlirenyuan { get; set; }
        public string J21guyouyewu { get; set; }
        public string J22xintuoyewu { get; set; }
        public string J23qita { get; set; }
        public string J24heji { get; set; }
        public string J25_30yixia { get; set; }
        public string J26_30_39yixia { get; set; }
        public string J27_40yishang { get; set; }
        //股东2018年  持股比例 

        public string J32 { get; set; }
        public string K32 { get; set; }
        public string J33 { get; set; }
        public string K33 { get; set; }
        public string J34 { get; set; }
        public string K34 { get; set; }
        public string J35 { get; set; }
        public string K35 { get; set; }
        public string J36 { get; set; }
        public string K36 { get; set; }
        public string J37 { get; set; }
        public string K37 { get; set; }
        public string J38 { get; set; }
        public string K38 { get; set; }
        public string J39 { get; set; }
        public string K39 { get; set; }
        public string J40 { get; set; }
        public string K40 { get; set; }
        public string J41 { get; set; }
        public string K41 { get; set; }
        // 纳入合并范围的子公司 	 注册地 	 注册资本万元 	 业务性质 	 持股比例 

        public string M32 { get; set; }
        public string N32 { get; set; }
        public string O32 { get; set; }
        public string P32 { get; set; }
        public string Q32 { get; set; }
        public string M33 { get; set; }
        public string N33 { get; set; }
        public string O33 { get; set; }
        public string P33 { get; set; }
        public string Q33 { get; set; }
        public string M34 { get; set; }
        public string N34 { get; set; }
        public string O34 { get; set; }
        public string P34 { get; set; }
        public string Q34 { get; set; }
        public string M35 { get; set; }
        public string N35 { get; set; }
        public string O35 { get; set; }
        public string P35 { get; set; }
        public string Q35 { get; set; }
        public string M36{ get; set; }
        public string N36 { get; set; }
        public string O36 { get; set; }
        public string P36 { get; set; }
        public string Q36 { get; set; }
        public string M37 { get; set; }
        public string N37 { get; set; }
        public string O37 { get; set; }
        public string P37 { get; set; }
        public string Q37 { get; set; }
        public string M38 { get; set; }
        public string N38 { get; set; }
        public string O38 { get; set; }
        public string P38 { get; set; }
        public string Q38 { get; set; }
        public string M39 { get; set; }
        public string N39 { get; set; }
        public string O39 { get; set; }
        public string P39 { get; set; }
        public string Q39 { get; set; }
        public string M40 { get; set; }
        public string N40 { get; set; }
        public string O40 { get; set; }
        public string P40 { get; set; }
        public string Q40 { get; set; }
    

        public string WorksheetName { get; set; }
    }

    public class clsFenbiaoInfo
    {
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }
        public string E { get; set; }
        public string F { get; set; }
        public string G { get; set; }
        public string H { get; set; }
        public string I { get; set; }
        public string J { get; set; }
        public string K { get; set; }
        public string L { get; set; }
        public string M { get; set; }
        public string N { get; set; }
        public string O { get; set; }
        public string P { get; set; }
        public string Q { get; set; }
        public string R { get; set; }
        public string S { get; set; }
        public string T { get; set; }
        public string U { get; set; }
        public string V { get; set; }
        public string W { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
        public string Z { get; set; }
        public string AA { get; set; }
        public string AB { get; set; }
        public string AC { get; set; }
        public string AD { get; set; }
        public string AE { get; set; }
        public string AF { get; set; }
        public string AG { get; set; }
        public string AH { get; set; }
        public string AI { get; set; }
        public string AJ { get; set; }
        public string AK { get; set; }
        public string AL { get; set; }
        public string AM { get; set; }
        public string AN { get; set; }
        public string AO { get; set; }
        public string AP { get; set; }
        public string AQ { get; set; }
        public string AR { get; set; }
        public string AS { get; set; }
        public string AT { get; set; }
        public string AU { get; set; }
        public string AV { get; set; }
        public string AW { get; set; }
        public string AX { get; set; }
        public string AY { get; set; }
        public string AZ { get; set; }
 
    }
    public class clsmoban_biaoInfo
    {
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }
        public string E { get; set; }
        public string F { get; set; }
        public string G { get; set; }
        public string H { get; set; }
        public string I { get; set; }
        public string J { get; set; }
        public string K { get; set; }
        public string L { get; set; }
        public string M { get; set; }
        public string N { get; set; }
        public string O { get; set; }
        public string P { get; set; }
        public string Q { get; set; }
        public string R { get; set; }
        public string S { get; set; }
        public string T { get; set; }
        public string U { get; set; }
        public string V { get; set; }
        public string W { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
        public string Z { get; set; }
        public string AA { get; set; }
        public string AB { get; set; }
        public string AC { get; set; }
        public string AD { get; set; }
        public string AE { get; set; }
        public string AF { get; set; }
        public string AG { get; set; }
        public string AH { get; set; }
        public string AI { get; set; }
        public string AJ { get; set; }
        public string AK { get; set; }
        public string AL { get; set; }
        public string AM { get; set; }
        public string AN { get; set; }
        public string AO { get; set; }
        public string AP { get; set; }
        public string AQ { get; set; }
        public string AR { get; set; }
        public string AS { get; set; }
        public string AT { get; set; }
        public string AU { get; set; }
        public string AV { get; set; }
        public string AW { get; set; }
        public string AX { get; set; }
        public string AY { get; set; }
        public string AZ { get; set; }

    }

    public class AddconnectGroup_info
    {
        public string _id { get; set; }//玩法种类
        public string name { get; set; }//玩法种类
        public string PCid { get; set; }//玩法种类

    }
    public class clsuserinfo
    {
        public string Order_id { get; set; }
        public string name { get; set; }
        public string password { get; set; }
        public string Btype { get; set; }
        public string denglushijian { get; set; }
        public string Createdate { get; set; }
        public string AdminIS { get; set; }
        public string jigoudaima { get; set; }
        public string pid { get; set; }

    }
    public class Addconnect_info
    {
        public string _id { get; set; }//玩法种类
        public string name { get; set; }//玩法种类
        public string mail { get; set; }//玩法种类
        public string address { get; set; }//玩法种类
        public string phone { get; set; }//玩法种类
        public string cmname { get; set; }//玩法种类
        public string weblink { get; set; }//玩法种类
        public string groupID { get; set; }//玩法种类
        public string PCid { get; set; }//玩法种类


    }
    public class FromGroup_info
    {
        public string _id { get; set; }//玩法种类
        public string name { get; set; }//玩法种类
        public string PCid { get; set; }//玩法种类

    }
    public class FromList_info
    {
        public string _id { get; set; }//玩法种类

        public string mail { get; set; }//玩法种类
        public string password { get; set; }//玩法种类
        public string mark { get; set; }//玩法种类
        public string groupID { get; set; }//玩法种类
        public string PCid { get; set; }//玩法种类



    }
    public class Template_info
    {
        public string _id { get; set; }

        public string subject { get; set; }
        public string body { get; set; }
        public string acc { get; set; }
        public string groupID { get; set; }
        public string PCid { get; set; }



    }
    public class AutoSend_info
    {
        public string _id { get; set; }

        public string zhuangtai { get; set; }
        public string zhuti { get; set; }
        public string neirong { get; set; }
        public string shoujianren { get; set; }
        public string fajianren { get; set; }
        public string kaishijian { get; set; }
        public string tingzhishijian { get; set; }
        public string jindu { get; set; }
        public string yaoqiuyueduhuizhi { get; set; }
        public string youxianji { get; set; }

    }
    public class Timer_info
    {
        public string _id { get; set; }

        public string time_start { get; set; }
        public string time_end { get; set; }
        public string TemplateID { get; set; }
        public string mail { get; set; }
        public string CCmail { get; set; }
        public string formto { get; set; }

        public string subject { get; set; }
        public string body { get; set; }
        public string acc { get; set; }
        public string groupID { get; set; }
        public string PCid { get; set; }
        public string status { get; set; }


    }


    public class softTime_info
    {
        public string _id { get; set; }//玩法种类

        public string starttime { get; set; }//玩法种类
        public string name { get; set; }//玩法种类
        public string endtime { get; set; }//玩法种类
        public string soft_name { get; set; }//玩法种类
        public string denglushijian { get; set; }//玩法种类


        public string password { get; set; }//玩法种类
        public string pid { get; set; }//玩法种类
        public string mark1 { get; set; }//玩法种类
        public string mark2 { get; set; }//玩法种类
        public string mark3 { get; set; }//玩法种类
        public string mark4 { get; set; }//玩法种类
        public string mark5 { get; set; }//玩法种类
    }
    public class clsQQquninfo
    {
        public string Order_id { get; set; }
        public string qun_name { get; set; }
        public string send_body { get; set; }
        public string is_timer { get; set; }
        public string send_time { get; set; }
        public string mark1 { get; set; }
        public string mark2 { get; set; }
        public string mark3 { get; set; }
        public string mark4 { get; set; }
        public string mark5 { get; set; }

    }
    public class clsalter_message
    {
        public string _id { get; set; }
        public string project_id { get; set; }
        public string project_name { get; set; }
        public string text { get; set; }
        public string mark1 { get; set; }
        public string mark2 { get; set; }
        public string mark3 { get; set; }
        public string mark4 { get; set; }
        public string mark5 { get; set; }

    }
    public class clsSendmailinfo
    {
        public string _id { get; set; }

        public string sendfrom { get; set; }
        public string sendto { get; set; }
        public string subject { get; set; }
        public string bodyinfo { get; set; }
        public string acc { get; set; }

        public string msg_tel { get; set; }

        public string host { get; set; }
        public string password { get; set; }

    
    }
}
