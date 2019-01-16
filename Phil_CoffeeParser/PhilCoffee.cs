using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using SpssLib.SpssDataset;
using SpssLib.DataReader;
using System.Data.SqlClient;

namespace Phil_CoffeeParser
{
    class PhilCoffee
    {
        static void Main(string[] args)
        {
            int SurveyID = 600531;
            //string SurveyPeriod = "2014-09-30";//survey period
            string SurveyPeriod = "2018-05-31";
            string Country = "Philippines";//survey country
            Phil_insertion iobj = new Phil_insertion();
            string questions = "Respondent_ID,WeightF,Gender,D1_MS,AREAS,S2_Age_Group,D3_EDUC,D4_WS,Sec,Dip,S5_01,S5_02,S5_03,S5_04,S5_05,S5_06,S5_07,S5_08,S5_09,S5_10,S5_11,S5_12,S6_01,S6_02,S6_03,S6_04,S6_05,S6_06,S6_07,S6_08,S6_09,S6_10,S6_11,S6_12,S7_01,S7_02,S7_03,S7_04,S7_05,S7_06,S7_07,S7_08,S7_09,S7_10,S7_11,S7_12,S8_01,S8_02,S8_03,S8_04,S8_05,S8_06,S8_07,S8_08,S8_09,S8_10,S8_11,S8_12,B1_tom_brand,B2_others_brand_51,B2_others_brand_73,B2_others_brand_88,B2_others_brand_75,B2_others_brand_21,B2_others_brand_23,B2_others_brand_50,B2_others_brand_54,B2_others_brand_310,B2_others_brand_312,B2_others_brand_711,B2_others_brand_712,B2_others_brand_807,B2_others_brand_1003,B4_aided_brand_51,B4_aided_brand_73,B4_aided_brand_88,B4_aided_brand_75,B4_aided_brand_21,B4_aided_brand_23,B4_aided_brand_50,B4_aided_brand_54,B4_aided_brand_310,B4_aided_brand_312,B4_aided_brand_711,B4_aided_brand_712,B4_aided_brand_807,B4_aided_brand_1003,B3_tom_ad,B3_others_ad_51,B3_others_ad_73,B3_others_ad_88,B3_others_ad_75,B3_others_ad_21,B3_others_ad_23,B3_others_ad_50,B3_others_ad_54,B3_others_ad_310,B3_others_ad_312,B3_others_ad_711,B3_others_ad_712,B3_others_ad_807,B3_others_ad_1003,B3_1_aided_ad_51,B3_1_aided_ad_73,B3_1_aided_ad_88,B3_1_aided_ad_75,B3_1_aided_ad_21,B3_1_aided_ad_23,B3_1_aided_ad_50,B3_1_aided_ad_54,B3_1_aided_ad_310,B3_1_aided_ad_312,B3_1_aided_ad_711,B3_1_aided_ad_712,B3_1_aided_ad_807,B3_1_aided_ad_1003,S9_01,S9_02,S9_03,S9_04,S9_05,S9_06,S9_07,Grid_S10_1_S10,Grid_S10_2_S10,Grid_S10_3_S10,Grid_S10_4_S10,Grid_S10_6_S10,Grid_S10_7_S10,Grid_S10_5_S10,Grid_S11_1_S11,Grid_S11_2_S11,Grid_S11_3_S11,Grid_S11_4_S11,Grid_S11_6_S11,Grid_S11_7_S11,Grid_S11_5_S11,B6_ever_tried_51,B6_ever_tried_73,B6_ever_tried_88,B6_ever_tried_75,B6_ever_tried_21,B6_ever_tried_23,B6_ever_tried_50,B6_ever_tried_54,B6_ever_tried_310,B6_ever_tried_312,B6_ever_tried_711,B6_ever_tried_712,B6_ever_tried_807,B6_ever_tried_1003,Grid_B7_51_B7,Grid_B7_73_B7,Grid_B7_88_B7,Grid_B7_75_B7,Grid_B7_21_B7,Grid_B7_23_B7,Grid_B7_50_B7,Grid_B7_54_B7,Grid_B7_310_B7,Grid_B7_312_B7,Grid_B7_711_B7,Grid_B7_712_B7,Grid_B7_807_B7,Grid_B7_1003_B7,B10_51,B10_73,B10_88,B10_75,B10_21,B10_23,B10_50,B10_54,B10_310,B10_312,B10_711,B10_712,B10_807,B10_1003,B11,B13,B13a,B16a,B16b,B1_TOM_BRAND_NET,B2_others_Brand_NET_3,B2_others_Brand_NET_6,B2_others_Brand_NET_7,B4_aided_Brand_NET_3,B4_aided_Brand_NET_6,B4_aided_Brand_NET_7,B3_tom_ad_BRAND_NET,B3_others_ad_BRAND_NET_3,B3_others_ad_BRAND_NET_6,B3_others_ad_BRAND_NET_7,B3_1_aided_ad_Brand_NET_3,B3_1_aided_ad_Brand_NET_6,B3_1_aided_ad_Brand_NET_7,B6_ever_tried_Brand_NET_3,B6_ever_tried_Brand_NET_6,B6_ever_tried_Brand_NET_7,B10_Brand_NET_3,B10_Brand_NET_6,B10_Brand_NET_7,B11_Brand_NET,B13_Brand_NET,B13a_Brand_NET,B16b_Brand_NET,D1b21_1,D1b21_4,D1b21_6,D1b21_14,D1b21_18,D1b21_31,D1b21_69,D1b21_71,D1b21_72,D1b21_75,D1b21_76,D1b21_78,D1b23_1,D1b23_4,D1b23_6,D1b23_14,D1b23_18,D1b23_31,D1b23_69,D1b23_71,D1b23_72,D1b23_75,D1b23_76,D1b23_78,D1b50_1,D1b50_4,D1b50_6,D1b50_14,D1b50_18,D1b50_31,D1b50_69,D1b50_71,D1b50_72,D1b50_75,D1b50_76,D1b50_78,D1b51_1,D1b51_4,D1b51_6,D1b51_14,D1b51_18,D1b51_31,D1b51_69,D1b51_71,D1b51_72,D1b51_75,D1b51_76,D1b51_78,D1b54_1,D1b54_4,D1b54_6,D1b54_14,D1b54_18,D1b54_31,D1b54_69,D1b54_71,D1b54_72,D1b54_75,D1b54_76,D1b54_78,D1b73_1,D1b73_4,D1b73_6,D1b73_14,D1b73_18,D1b73_31,D1b73_69,D1b73_71,D1b73_72,D1b73_75,D1b73_76,D1b73_78,D1b75_1,D1b75_4,D1b75_6,D1b75_14,D1b75_18,D1b75_31,D1b75_69,D1b75_71,D1b75_72,D1b75_75,D1b75_76,D1b75_78,D1b88_1,D1b88_4,D1b88_6,D1b88_14,D1b88_18,D1b88_31,D1b88_69,D1b88_71,D1b88_72,D1b88_75,D1b88_76,D1b88_78";// dashboard variable value
            //string questions = "AREAS";
            string[] spss_variable_name = questions.Split(',');
            SpssReader spssDataset;
            using (FileStream fileStream = new FileStream(@"E:\Clue-Box Work\2018\JUN-2018\PH-COFF MAY-2018\REC\PHC-MAY-2018.sav", FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10, FileOptions.SequentialScan))
            {
                spssDataset = new SpssReader(fileStream); // Create the reader, this will read the file header
                foreach (string spss_v in spss_variable_name)
                {
                    foreach (var variable in spssDataset.Variables)  // Iterate through all the varaibles
                    {
                        if (variable.Name.Equals(spss_v))
                        {
                            foreach (KeyValuePair<double, string> label in variable.ValueLabels)
                            {
                                string VARIABLE_NAME = spss_v;
                                string VARIABLE_NAME_SUB_NAME = variable.Name;
                                string VARIABLE_ID = label.Key.ToString();
                                string VARIABLE_VALUE = label.Value;
                                string VARIABLE_NAME_QUESTION = variable.Label;
                                string BASE_VARIABLE_NAME = variable.Name;
                               // iobj.insert_Dashboard_variable_values(VARIABLE_NAME, VARIABLE_NAME_SUB_NAME, VARIABLE_ID, VARIABLE_VALUE, VARIABLE_NAME_QUESTION, SurveyID, Country, BASE_VARIABLE_NAME, SurveyPeriod);
                            }
                        }

                    }
                }
                foreach (var record in spssDataset.Records)
                {
                    string UserID = null;
                    string variable_name;
                    string u_id = null;
                    decimal Weight = 0;
                    string AttendedOn = "-- Not Available --";
                    string Gender = "-- Not Available --";
                    string MaritalStatus = "-- Not Available --";
                    string Location = "-- Not Available --";
                    string AgeGroup = "-- Not Available --";
                    string Edulevel = "-- Not Available --";
                    string Occupation = "-- Not Available --";
                    string Ses = "-- Not Available --";
                    string Period = "-- Not Available --";
                    string BevConP6MBotWat = "-- Not Available --";
                    string BevConP6MCoffee = "-- Not Available --";
                    string BevConP6MFrjuice = "-- Not Available --";
                    string BevConP6MIcTea = "-- Not Available --";
                    string BevConP6MSoftdrinks = "-- Not Available --";
                    string BevConP6MEnDrinks = "-- Not Available --";
                    string BevConP6MChocDrink = "-- Not Available --";
                    string BevConP6MCerDrink = "-- Not Available --";
                    string BevConP6MMilk = "-- Not Available --";
                    string BevConP6MHotCoffee = "-- Not Available --";
                    string BevConP6MColdCoffee = "-- Not Available --";
                    string BevConP6MNoneAb = "-- Not Available --";
                    string BevConP3MBotWat = "-- Not Available --";
                    string BevConP3MCoffee = "-- Not Available --";
                    string BevConP3MFrjuice = "-- Not Available --";
                    string BevConP3MIcTea = "-- Not Available --";
                    string BevConP3MSoftdrinks = "-- Not Available --";
                    string BevConP3MEnDrinks = "-- Not Available --";
                    string BevConP3MChocDrink = "-- Not Available --";
                    string BevConP3MCerDrink = "-- Not Available --";
                    string BevConP3MMilk = "-- Not Available --";
                    string BevConP3MHotCoffee = "-- Not Available --";
                    string BevConP3MColdCoffee = "-- Not Available --";
                    string BevConP3MNoneAb = "-- Not Available --";
                    string BevConP1MBotWat = "-- Not Available --";
                    string BevConP1MCoffee = "-- Not Available --";
                    string BevConP1MFrjuice = "-- Not Available --";
                    string BevConP1MIcTea = "-- Not Available --";
                    string BevConP1MSoftdrinks = "-- Not Available --";
                    string BevConP1MEnDrinks = "-- Not Available --";
                    string BevConP1MChocDrink = "-- Not Available --";
                    string BevConP1MCerDrink = "-- Not Available --";
                    string BevConP1MMilk = "-- Not Available --";
                    string BevConP1MHotCoffee = "-- Not Available --";
                    string BevConP1MColdCoffee = "-- Not Available --";
                    string BevConP1MNoneAb = "-- Not Available --";
                    string BevConP1WBotWat = "-- Not Available --";
                    string BevConP1WCoffee = "-- Not Available --";
                    string BevConP1WFrjuice = "-- Not Available --";
                    string BevConP1WIcTea = "-- Not Available --";
                    string BevConP1WSoftdrinks = "-- Not Available --";
                    string BevConP1WEnDrinks = "-- Not Available --";
                    string BevConP1WChocDrink = "-- Not Available --";
                    string BevConP1WCerDrink = "-- Not Available --";
                    string BevConP1WMilk = "-- Not Available --";
                    string BevConP1WHotCoffee = "-- Not Available --";
                    string BevConP1WColdCoffee = "-- Not Available --";
                    string BevConP1WNoneAb = "-- Not Available --";
                    string BrTom = "-- Not Available --";
                    string BrSpontKOPIKOBROWN = "-- Not Available --";
                    string BrSpontN_CreamyLatte = "-- Not Available --";
                    string BrSpontNCW = "-- Not Available --";
                    string BrSpontN_Brn_Creamy = "-- Not Available --";
                    string BrSpontGTW = "-- Not Available --";
                    string BrSpontGTO_TWINPK = "-- Not Available --";
                    string BrSpontKOPIKOASTIG = "-- Not Available --";
                    string BrSpontKopikoCafBlnca = "-- Not Available --";
                    string BrSpontGTW_TwinPk = "-- Not Available --";
                    string BrSpontNCW_TwinPk = "-- Not Available --";
                    string BrSpontKopikoAstig_TwPk = "-- Not Available --";
                    string BrSpontKopikoBrn_TwPk = "-- Not Available --";
                    string BrSpontKopikoBlca_TwPk = "-- Not Available --";
                    string BrSpontN_Bl_Brew_Br = "-- Not Available --";
                    string BrAidKOPIKOBROWN = "-- Not Available --";
                    string BrAidN_CreamyLatte = "-- Not Available --";
                    string BrAidNCW = "-- Not Available --";
                    string BrAidN_Brn_Creamy = "-- Not Available --";
                    string BrAidGTW = "-- Not Available --";
                    string BrAidGTO_TWINPK = "-- Not Available --";
                    string BrAidKOPIKOASTIG = "-- Not Available --";
                    string BrAidKopikoCafBlnca = "-- Not Available --";
                    string BrAidGTW_TwinPk = "-- Not Available --";
                    string BrAidNCW_TwinPk = "-- Not Available --";
                    string BrAidKopikoAstig_TwPk = "-- Not Available --";
                    string BrAidKopikoBrn_TwPk = "-- Not Available --";
                    string BrAidKopikoBlca_TwPk = "-- Not Available --";
                    string BrAidN_Bl_Brew_Br = "-- Not Available --";
                    string AdTom = "-- Not Available --";
                    string AdSpontKOPIKOBROWN = "-- Not Available --";
                    string AdSpontN_CreamyLatte = "-- Not Available --";
                    string AdSpontNCW = "-- Not Available --";
                    string AdSpontN_Brn_Creamy = "-- Not Available --";
                    string AdSpontGTW = "-- Not Available --";
                    string AdSpontGTO_TWINPK = "-- Not Available --";
                    string AdSpontKOPIKOASTIG = "-- Not Available --";
                    string AdSpontKopikoCafBlnca = "-- Not Available --";
                    string AdSpontGTW_TwinPk = "-- Not Available --";
                    string AdSpontNCW_TwinPk = "-- Not Available --";
                    string AdSpontKopikoAstig_TwPk = "-- Not Available --";
                    string AdSpontKopikoBrn_TwPk = "-- Not Available --";
                    string AdSpontKopikoBlca_TwPk = "-- Not Available --";
                    string AdSpontN_Bl_Brew_Br = "-- Not Available --";
                    string AdAidKOPIKOBROWN = "-- Not Available --";
                    string AdAidN_CreamyLatte = "-- Not Available --";
                    string AdAidNCW = "-- Not Available --";
                    string AdAidN_Brn_Creamy = "-- Not Available --";
                    string AdAidGTW = "-- Not Available --";
                    string AdAidGTO_TWINPK = "-- Not Available --";
                    string AdAidKOPIKOASTIG = "-- Not Available --";
                    string AdAidKopikoCafBlnca = "-- Not Available --";
                    string AdAidGTW_TwinPk = "-- Not Available --";
                    string AdAidNCW_TwinPk = "-- Not Available --";
                    string AdAidKopikoAstig_TwPk = "-- Not Available --";
                    string AdAidKopikoBrn_TwPk = "-- Not Available --";
                    string AdAidKopikoBlca_TwPk = "-- Not Available --";
                    string AdAidN_Bl_Brew_Br = "-- Not Available --";
                    string CofConP1MRoasted = "-- Not Available --";
                    string CofConP1MVend = "-- Not Available --";
                    string CofConP1MIntCoffee = "-- Not Available --";
                    string CofConP1MIntPowCoff = "-- Not Available --";
                    string CofConP1MRTDCoff = "-- Not Available --";
                    string CofConP1MCoffCaf = "-- Not Available --";
                    string CofConP1MBlackCoff = "-- Not Available --";
                    string CofCon1WRoasted = "-- Not Available --";
                    string CofCon1WVend = "-- Not Available --";
                    string CofCon1WIntCoffee = "-- Not Available --";
                    string CofCon1WIntPowCoff = "-- Not Available --";
                    string CofCon1WRTDCoff = "-- Not Available --";
                    string CofCon1WCoffCaf = "-- Not Available --";
                    string CofCon1WBlackCoff = "-- Not Available --";
                    string CofCon1DRoasted = "-- Not Available --";
                    string CofCon1DVend = "-- Not Available --";
                    string CofCon1DIntCoffee = "-- Not Available --";
                    string CofCon1DIntPowCoff = "-- Not Available --";
                    string CofCon1DRTDCoff = "-- Not Available --";
                    string CofCon1DCoffCaf = "-- Not Available --";
                    string CofCon1DBlackCoff = "-- Not Available --";
                    string EvTriKOPIKOBROWN = "-- Not Available --";
                    string EvTriN_CreamyLatte = "-- Not Available --";
                    string EvTriNCW = "-- Not Available --";
                    string EvTriN_Brn_Creamy = "-- Not Available --";
                    string EvTriGTW = "-- Not Available --";
                    string EvTriGTO_TWINPK = "-- Not Available --";
                    string EvTriKOPIKOASTIG = "-- Not Available --";
                    string EvTriKopikoCafBlnca = "-- Not Available --";
                    string EvTriGTW_TwinPk = "-- Not Available --";
                    string EvTriNCW_TwinPk = "-- Not Available --";
                    string EvTriKopikoAstig_TwPk = "-- Not Available --";
                    string EvTriKopikoBrn_TwPk = "-- Not Available --";
                    string EvTriKopikoBlca_TwPk = "-- Not Available --";
                    string EvTriN_Bl_Brew_Br = "-- Not Available --";
                    string LaConsKOPIKOBROWN = "-- Not Available --";
                    string LaConsN_CreamyLatte = "-- Not Available --";
                    string LaConsNCW = "-- Not Available --";
                    string LaConsN_Brn_Creamy = "-- Not Available --";
                    string LaConsGTW = "-- Not Available --";
                    string LaConsGTO_TWINPK = "-- Not Available --";
                    string LaConsKOPIKOASTIG = "-- Not Available --";
                    string LaConsKopikoCafBlnca = "-- Not Available --";
                    string LaConsGTW_TwinPk = "-- Not Available --";
                    string LaConsNCW_TwinPk = "-- Not Available --";
                    string LaConsKopikoAstig_TwPk = "-- Not Available --";
                    string LaConsKopikoBrn_TwPk = "-- Not Available --";
                    string LaConsKopikoBlca_TwPk = "-- Not Available --";
                    string LaConsN_Bl_Brew_Br = "-- Not Available --";
                    string RegP3MKOPIKOBROWN = "-- Not Available --";
                    string RegP3MN_CreamyLatte = "-- Not Available --";
                    string RegP3MNCW = "-- Not Available --";
                    string RegP3MN_Brn_Creamy = "-- Not Available --";
                    string RegP3MGTW = "-- Not Available --";
                    string RegP3MGTO_TWINPK = "-- Not Available --";
                    string RegP3MKOPIKOASTIG = "-- Not Available --";
                    string RegP3MKopikoCafBlnca = "-- Not Available --";
                    string RegP3MGTW_TwinPk = "-- Not Available --";
                    string RegP3MNCW_TwinPk = "-- Not Available --";
                    string RegP3MKopikoAstig_TwPk = "-- Not Available --";
                    string RegP3MKopikoBrn_TwPk = "-- Not Available --";
                    string RegP3MKopikoBlca_TwPk = "-- Not Available --";
                    string RegP3MN_Bl_Brew_Br = "-- Not Available --";
                    string Bumo = "-- Not Available --";
                    string PreBumo = "-- Not Available --";
                    string FavBrand = "-- Not Available --";
                    string IntSwBr_6M = "-- Not Available --";
                    string BrLikeSw = "-- Not Available --";
                    string NetBrTom = "-- Not Available --";
                    string NetBrSpontGreat_Taste = "-- Not Available --";
                    string NetBrSpontKopiko = "-- Not Available --";
                    string NetBrSpontNescafe = "-- Not Available --";
                    string NetBrAidGreat_Taste = "-- Not Available --";
                    string NetBrAidKopiko = "-- Not Available --";
                    string NetBrAidNescafe = "-- Not Available --";
                    string NetAdTom = "-- Not Available --";
                    string NetAdSpontGreat_Taste = "-- Not Available --";
                    string NetAdSpontKopiko = "-- Not Available --";
                    string NetAdSpontNescafe = "-- Not Available --";
                    string NetAdAidGreat_Taste = "-- Not Available --";
                    string NetAdAidKopiko = "-- Not Available --";
                    string NetAdAidNescafe = "-- Not Available --";
                    string NetEvTriGreat_Taste = "-- Not Available --";
                    string NetEvTriKopiko = "-- Not Available --";
                    string NetEvTriNescafe = "-- Not Available --";
                    string NetRegP3MGreat_Taste = "-- Not Available --";
                    string NetRegP3MKopiko = "-- Not Available --";
                    string NetRegP3MNescafe = "-- Not Available --";
                    string NetBumo = "-- Not Available --";
                    string NetPreBumo = "-- Not Available --";
                    string NetFavBrand = "-- Not Available --";
                    string NetBrLikeSw = "-- Not Available --";
                    string D1b21_1 = "-- Not Available --";
                    string D1b21_4 = "-- Not Available --";
                    string D1b21_6 = "-- Not Available --";
                    string D1b21_14 = "-- Not Available --";
                    string D1b21_18 = "-- Not Available --";
                    string D1b21_31 = "-- Not Available --";
                    string D1b21_69 = "-- Not Available --";
                    string D1b21_71 = "-- Not Available --";
                    string D1b21_72 = "-- Not Available --";
                    string D1b21_75 = "-- Not Available --";
                    string D1b21_76 = "-- Not Available --";
                    string D1b21_78 = "-- Not Available --";
                    string D1b23_1 = "-- Not Available --";
                    string D1b23_4 = "-- Not Available --";
                    string D1b23_6 = "-- Not Available --";
                    string D1b23_14 = "-- Not Available --";
                    string D1b23_18 = "-- Not Available --";
                    string D1b23_31 = "-- Not Available --";
                    string D1b23_69 = "-- Not Available --";
                    string D1b23_71 = "-- Not Available --";
                    string D1b23_72 = "-- Not Available --";
                    string D1b23_75 = "-- Not Available --";
                    string D1b23_76 = "-- Not Available --";
                    string D1b23_78 = "-- Not Available --";
                    string D1b50_1 = "-- Not Available --";
                    string D1b50_4 = "-- Not Available --";
                    string D1b50_6 = "-- Not Available --";
                    string D1b50_14 = "-- Not Available --";
                    string D1b50_18 = "-- Not Available --";
                    string D1b50_31 = "-- Not Available --";
                    string D1b50_69 = "-- Not Available --";
                    string D1b50_71 = "-- Not Available --";
                    string D1b50_72 = "-- Not Available --";
                    string D1b50_75 = "-- Not Available --";
                    string D1b50_76 = "-- Not Available --";
                    string D1b50_78 = "-- Not Available --";
                    string D1b51_1 = "-- Not Available --";
                    string D1b51_4 = "-- Not Available --";
                    string D1b51_6 = "-- Not Available --";
                    string D1b51_14 = "-- Not Available --";
                    string D1b51_18 = "-- Not Available --";
                    string D1b51_31 = "-- Not Available --";
                    string D1b51_69 = "-- Not Available --";
                    string D1b51_71 = "-- Not Available --";
                    string D1b51_72 = "-- Not Available --";
                    string D1b51_75 = "-- Not Available --";
                    string D1b51_76 = "-- Not Available --";
                    string D1b51_78 = "-- Not Available --";
                    string D1b54_1 = "-- Not Available --";
                    string D1b54_4 = "-- Not Available --";
                    string D1b54_6 = "-- Not Available --";
                    string D1b54_14 = "-- Not Available --";
                    string D1b54_18 = "-- Not Available --";
                    string D1b54_31 = "-- Not Available --";
                    string D1b54_69 = "-- Not Available --";
                    string D1b54_71 = "-- Not Available --";
                    string D1b54_72 = "-- Not Available --";
                    string D1b54_75 = "-- Not Available --";
                    string D1b54_76 = "-- Not Available --";
                    string D1b54_78 = "-- Not Available --";
                    string D1b73_1 = "-- Not Available --";
                    string D1b73_4 = "-- Not Available --";
                    string D1b73_6 = "-- Not Available --";
                    string D1b73_14 = "-- Not Available --";
                    string D1b73_18 = "-- Not Available --";
                    string D1b73_31 = "-- Not Available --";
                    string D1b73_69 = "-- Not Available --";
                    string D1b73_71 = "-- Not Available --";
                    string D1b73_72 = "-- Not Available --";
                    string D1b73_75 = "-- Not Available --";
                    string D1b73_76 = "-- Not Available --";
                    string D1b73_78 = "-- Not Available --";
                    string D1b75_1 = "-- Not Available --";
                    string D1b75_4 = "-- Not Available --";
                    string D1b75_6 = "-- Not Available --";
                    string D1b75_14 = "-- Not Available --";
                    string D1b75_18 = "-- Not Available --";
                    string D1b75_31 = "-- Not Available --";
                    string D1b75_69 = "-- Not Available --";
                    string D1b75_71 = "-- Not Available --";
                    string D1b75_72 = "-- Not Available --";
                    string D1b75_75 = "-- Not Available --";
                    string D1b75_76 = "-- Not Available --";
                    string D1b75_78 = "-- Not Available --";
                    string D1b88_1 = "-- Not Available --";
                    string D1b88_4 = "-- Not Available --";
                    string D1b88_6 = "-- Not Available --";
                    string D1b88_14 = "-- Not Available --";
                    string D1b88_18 = "-- Not Available --";
                    string D1b88_31 = "-- Not Available --";
                    string D1b88_69 = "-- Not Available --";
                    string D1b88_71 = "-- Not Available --";
                    string D1b88_72 = "-- Not Available --";
                    string D1b88_75 = "-- Not Available --";
                    string D1b88_76 = "-- Not Available --";
                    string D1b88_78 = "-- Not Available --";


                    foreach (var variable in spssDataset.Variables)
                    {
                        foreach (string spss_v in spss_variable_name)
                        {
                            if (variable.Name.Equals(spss_v))
                            {
                                variable_name = variable.Name;
                                switch (variable_name)
                                {
                                    case "Respondent_ID":
                                        {
                                            u_id = Convert.ToString(record.GetValue(variable));
                                            UserID = find_UserID(SurveyID, SurveyPeriod, u_id);
                                            //userID = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "WeightF":
                                        {
                                            Weight = Convert.ToDecimal(record.GetValue(variable));
                                            break;
                                        }
                                    case "Gender":
                                        {
                                            Gender = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1_MS":
                                        {
                                            MaritalStatus = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "AREAS":
                                        {
                                            Location = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S2_Age_Group":
                                        {
                                            AgeGroup = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D3_EDUC":
                                        {
                                            Edulevel = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D4_WS":
                                        {
                                            Occupation = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Sec":
                                        {
                                            Ses = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Dip":
                                        {
                                            Period = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_01":
                                        {
                                            BevConP6MBotWat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_02":
                                        {
                                            BevConP6MCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_03":
                                        {
                                            BevConP6MFrjuice = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_04":
                                        {
                                            BevConP6MIcTea = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_05":
                                        {
                                            BevConP6MSoftdrinks = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_06":
                                        {
                                            BevConP6MEnDrinks = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_07":
                                        {
                                            BevConP6MChocDrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_08":
                                        {
                                            BevConP6MCerDrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_09":
                                        {
                                            BevConP6MMilk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_10":
                                        {
                                            BevConP6MHotCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_11":
                                        {
                                            BevConP6MColdCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S5_12":
                                        {
                                            BevConP6MNoneAb = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_01":
                                        {
                                            BevConP3MBotWat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_02":
                                        {
                                            BevConP3MCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_03":
                                        {
                                            BevConP3MFrjuice = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_04":
                                        {
                                            BevConP3MIcTea = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_05":
                                        {
                                            BevConP3MSoftdrinks = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_06":
                                        {
                                            BevConP3MEnDrinks = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_07":
                                        {
                                            BevConP3MChocDrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_08":
                                        {
                                            BevConP3MCerDrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_09":
                                        {
                                            BevConP3MMilk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_10":
                                        {
                                            BevConP3MHotCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_11":
                                        {
                                            BevConP3MColdCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S6_12":
                                        {
                                            BevConP3MNoneAb = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_01":
                                        {
                                            BevConP1MBotWat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_02":
                                        {
                                            BevConP1MCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_03":
                                        {
                                            BevConP1MFrjuice = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_04":
                                        {
                                            BevConP1MIcTea = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_05":
                                        {
                                            BevConP1MSoftdrinks = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_06":
                                        {
                                            BevConP1MEnDrinks = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_07":
                                        {
                                            BevConP1MChocDrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_08":
                                        {
                                            BevConP1MCerDrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_09":
                                        {
                                            BevConP1MMilk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_10":
                                        {
                                            BevConP1MHotCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_11":
                                        {
                                            BevConP1MColdCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S7_12":
                                        {
                                            BevConP1MNoneAb = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_01":
                                        {
                                            BevConP1WBotWat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_02":
                                        {
                                            BevConP1WCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_03":
                                        {
                                            BevConP1WFrjuice = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_04":
                                        {
                                            BevConP1WIcTea = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_05":
                                        {
                                            BevConP1WSoftdrinks = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_06":
                                        {
                                            BevConP1WEnDrinks = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_07":
                                        {
                                            BevConP1WChocDrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_08":
                                        {
                                            BevConP1WCerDrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_09":
                                        {
                                            BevConP1WMilk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_10":
                                        {
                                            BevConP1WHotCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_11":
                                        {
                                            BevConP1WColdCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S8_12":
                                        {
                                            BevConP1WNoneAb = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B1_tom_brand":
                                        {
                                            BrTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_51":
                                        {
                                            BrSpontKOPIKOBROWN = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_73":
                                        {
                                            BrSpontN_CreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_88":
                                        {
                                            BrSpontNCW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_75":
                                        {
                                            BrSpontN_Brn_Creamy = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_21":
                                        {
                                            BrSpontGTW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_23":
                                        {
                                            BrSpontGTO_TWINPK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_50":
                                        {
                                            BrSpontKOPIKOASTIG = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_54":
                                        {
                                            BrSpontKopikoCafBlnca = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_310":
                                        {
                                            BrSpontGTW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_312":
                                        {
                                            BrSpontNCW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_711":
                                        {
                                            BrSpontKopikoAstig_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_712":
                                        {
                                            BrSpontKopikoBrn_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_807":
                                        {
                                            BrSpontKopikoBlca_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_brand_1003":
                                        {
                                            BrSpontN_Bl_Brew_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_51":
                                        {
                                            BrAidKOPIKOBROWN = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_73":
                                        {
                                            BrAidN_CreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_88":
                                        {
                                            BrAidNCW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_75":
                                        {
                                            BrAidN_Brn_Creamy = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_21":
                                        {
                                            BrAidGTW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_23":
                                        {
                                            BrAidGTO_TWINPK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_50":
                                        {
                                            BrAidKOPIKOASTIG = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_54":
                                        {
                                            BrAidKopikoCafBlnca = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_310":
                                        {
                                            BrAidGTW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_312":
                                        {
                                            BrAidNCW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_711":
                                        {
                                            BrAidKopikoAstig_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_712":
                                        {
                                            BrAidKopikoBrn_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_807":
                                        {
                                            BrAidKopikoBlca_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_brand_1003":
                                        {
                                            BrAidN_Bl_Brew_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_tom_ad":
                                        {
                                            AdTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_51":
                                        {
                                            AdSpontKOPIKOBROWN = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_73":
                                        {
                                            AdSpontN_CreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_88":
                                        {
                                            AdSpontNCW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_75":
                                        {
                                            AdSpontN_Brn_Creamy = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_21":
                                        {
                                            AdSpontGTW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_23":
                                        {
                                            AdSpontGTO_TWINPK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_50":
                                        {
                                            AdSpontKOPIKOASTIG = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_54":
                                        {
                                            AdSpontKopikoCafBlnca = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_310":
                                        {
                                            AdSpontGTW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_312":
                                        {
                                            AdSpontNCW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_711":
                                        {
                                            AdSpontKopikoAstig_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_712":
                                        {
                                            AdSpontKopikoBrn_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_807":
                                        {
                                            AdSpontKopikoBlca_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_1003":
                                        {
                                            AdSpontN_Bl_Brew_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_51":
                                        {
                                            AdAidKOPIKOBROWN = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_73":
                                        {
                                            AdAidN_CreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_88":
                                        {
                                            AdAidNCW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_75":
                                        {
                                            AdAidN_Brn_Creamy = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_21":
                                        {
                                            AdAidGTW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_23":
                                        {
                                            AdAidGTO_TWINPK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_50":
                                        {
                                            AdAidKOPIKOASTIG = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_54":
                                        {
                                            AdAidKopikoCafBlnca = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_310":
                                        {
                                            AdAidGTW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_312":
                                        {
                                            AdAidNCW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_711":
                                        {
                                            AdAidKopikoAstig_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_712":
                                        {
                                            AdAidKopikoBrn_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_807":
                                        {
                                            AdAidKopikoBlca_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_1003":
                                        {
                                            AdAidN_Bl_Brew_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S9_01":
                                        {
                                            CofConP1MRoasted = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S9_02":
                                        {
                                            CofConP1MVend = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S9_03":
                                        {
                                            CofConP1MIntCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S9_04":
                                        {
                                            CofConP1MIntPowCoff = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S9_05":
                                        {
                                            CofConP1MRTDCoff = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S9_06":
                                        {
                                            CofConP1MCoffCaf = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S9_07":
                                        {
                                            CofConP1MBlackCoff = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S10_1_S10":
                                        {
                                            CofCon1WRoasted = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S10_2_S10":
                                        {
                                            CofCon1WVend = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S10_3_S10":
                                        {
                                            CofCon1WIntCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S10_4_S10":
                                        {
                                            CofCon1WIntPowCoff = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S10_6_S10":
                                        {
                                            CofCon1WRTDCoff = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S10_7_S10":
                                        {
                                            CofCon1WCoffCaf = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S10_5_S10":
                                        {
                                            CofCon1WBlackCoff = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S11_1_S11":
                                        {
                                            CofCon1DRoasted = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S11_2_S11":
                                        {
                                            CofCon1DVend = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S11_3_S11":
                                        {
                                            CofCon1DIntCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S11_4_S11":
                                        {
                                            CofCon1DIntPowCoff = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S11_6_S11":
                                        {
                                            CofCon1DRTDCoff = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S11_7_S11":
                                        {
                                            CofCon1DCoffCaf = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_S11_5_S11":
                                        {
                                            CofCon1DBlackCoff = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_51":
                                        {
                                            EvTriKOPIKOBROWN = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_73":
                                        {
                                            EvTriN_CreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_88":
                                        {
                                            EvTriNCW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_75":
                                        {
                                            EvTriN_Brn_Creamy = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_21":
                                        {
                                            EvTriGTW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_23":
                                        {
                                            EvTriGTO_TWINPK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_50":
                                        {
                                            EvTriKOPIKOASTIG = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_54":
                                        {
                                            EvTriKopikoCafBlnca = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_310":
                                        {
                                            EvTriGTW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_312":
                                        {
                                            EvTriNCW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_711":
                                        {
                                            EvTriKopikoAstig_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_712":
                                        {
                                            EvTriKopikoBrn_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_807":
                                        {
                                            EvTriKopikoBlca_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_1003":
                                        {
                                            EvTriN_Bl_Brew_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_51_B7":
                                        {
                                            LaConsKOPIKOBROWN = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_73_B7":
                                        {
                                            LaConsN_CreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_88_B7":
                                        {
                                            LaConsNCW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_75_B7":
                                        {
                                            LaConsN_Brn_Creamy = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_21_B7":
                                        {
                                            LaConsGTW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_23_B7":
                                        {
                                            LaConsGTO_TWINPK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_50_B7":
                                        {
                                            LaConsKOPIKOASTIG = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_54_B7":
                                        {
                                            LaConsKopikoCafBlnca = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_310_B7":
                                        {
                                            LaConsGTW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_312_B7":
                                        {
                                            LaConsNCW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_711_B7":
                                        {
                                            LaConsKopikoAstig_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_712_B7":
                                        {
                                            LaConsKopikoBrn_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_807_B7":
                                        {
                                            LaConsKopikoBlca_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "Grid_B7_1003_B7":
                                        {
                                            LaConsN_Bl_Brew_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_51":
                                        {
                                            RegP3MKOPIKOBROWN = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_73":
                                        {
                                            RegP3MN_CreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_88":
                                        {
                                            RegP3MNCW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_75":
                                        {
                                            RegP3MN_Brn_Creamy = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_21":
                                        {
                                            RegP3MGTW = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_23":
                                        {
                                            RegP3MGTO_TWINPK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_50":
                                        {
                                            RegP3MKOPIKOASTIG = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_54":
                                        {
                                            RegP3MKopikoCafBlnca = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_310":
                                        {
                                            RegP3MGTW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_312":
                                        {
                                            RegP3MNCW_TwinPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_711":
                                        {
                                            RegP3MKopikoAstig_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_712":
                                        {
                                            RegP3MKopikoBrn_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_807":
                                        {
                                            RegP3MKopikoBlca_TwPk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_1003":
                                        {
                                            RegP3MN_Bl_Brew_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B11":
                                        {
                                            Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B13":
                                        {
                                            PreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B13a":
                                        {
                                            FavBrand = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B16a":
                                        {
                                            IntSwBr_6M = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B16b":
                                        {
                                            BrLikeSw = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B1_TOM_BRAND_NET":
                                        {
                                            NetBrTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_Brand_NET_3":
                                        {
                                            NetBrSpontGreat_Taste = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_Brand_NET_6":
                                        {
                                            NetBrSpontKopiko = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B2_others_Brand_NET_7":
                                        {
                                            NetBrSpontNescafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_Brand_NET_3":
                                        {
                                            NetBrAidGreat_Taste = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_Brand_NET_6":
                                        {
                                            NetBrAidKopiko = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B4_aided_Brand_NET_7":
                                        {
                                            NetBrAidNescafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_tom_ad_BRAND_NET":
                                        {
                                            NetAdTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_BRAND_NET_3":
                                        {
                                            NetAdSpontGreat_Taste = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_BRAND_NET_6":
                                        {
                                            NetAdSpontKopiko = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_others_ad_BRAND_NET_7":
                                        {
                                            NetAdSpontNescafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_Brand_NET_3":
                                        {
                                            NetAdAidGreat_Taste = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_Brand_NET_6":
                                        {
                                            NetAdAidKopiko = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B3_1_aided_ad_Brand_NET_7":
                                        {
                                            NetAdAidNescafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_Brand_NET_3":
                                        {
                                            NetEvTriGreat_Taste = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_Brand_NET_6":
                                        {
                                            NetEvTriKopiko = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B6_ever_tried_Brand_NET_7":
                                        {
                                            NetEvTriNescafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_Brand_NET_3":
                                        {
                                            NetRegP3MGreat_Taste = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_Brand_NET_6":
                                        {
                                            NetRegP3MKopiko = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B10_Brand_NET_7":
                                        {
                                            NetRegP3MNescafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B11_Brand_NET":
                                        {
                                            NetBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B13_Brand_NET":
                                        {
                                            NetPreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B13a_Brand_NET":
                                        {
                                            NetFavBrand = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "B16b_Brand_NET":
                                        {
                                            NetBrLikeSw = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_1":
                                        {
                                            D1b21_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_4":
                                        {
                                            D1b21_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_6":
                                        {
                                            D1b21_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_14":
                                        {
                                            D1b21_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_18":
                                        {
                                            D1b21_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_31":
                                        {
                                            D1b21_31 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_69":
                                        {
                                            D1b21_69 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_71":
                                        {
                                            D1b21_71 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_72":
                                        {
                                            D1b21_72 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_75":
                                        {
                                            D1b21_75 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_76":
                                        {
                                            D1b21_76 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b21_78":
                                        {
                                            D1b21_78 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_1":
                                        {
                                            D1b23_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_4":
                                        {
                                            D1b23_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_6":
                                        {
                                            D1b23_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_14":
                                        {
                                            D1b23_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_18":
                                        {
                                            D1b23_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_31":
                                        {
                                            D1b23_31 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_69":
                                        {
                                            D1b23_69 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_71":
                                        {
                                            D1b23_71 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_72":
                                        {
                                            D1b23_72 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_75":
                                        {
                                            D1b23_75 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_76":
                                        {
                                            D1b23_76 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b23_78":
                                        {
                                            D1b23_78 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_1":
                                        {
                                            D1b50_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_4":
                                        {
                                            D1b50_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_6":
                                        {
                                            D1b50_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_14":
                                        {
                                            D1b50_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_18":
                                        {
                                            D1b50_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_31":
                                        {
                                            D1b50_31 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_69":
                                        {
                                            D1b50_69 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_71":
                                        {
                                            D1b50_71 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_72":
                                        {
                                            D1b50_72 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_75":
                                        {
                                            D1b50_75 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_76":
                                        {
                                            D1b50_76 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b50_78":
                                        {
                                            D1b50_78 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_1":
                                        {
                                            D1b51_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_4":
                                        {
                                            D1b51_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_6":
                                        {
                                            D1b51_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_14":
                                        {
                                            D1b51_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_18":
                                        {
                                            D1b51_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_31":
                                        {
                                            D1b51_31 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_69":
                                        {
                                            D1b51_69 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_71":
                                        {
                                            D1b51_71 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_72":
                                        {
                                            D1b51_72 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_75":
                                        {
                                            D1b51_75 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_76":
                                        {
                                            D1b51_76 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b51_78":
                                        {
                                            D1b51_78 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_1":
                                        {
                                            D1b54_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_4":
                                        {
                                            D1b54_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_6":
                                        {
                                            D1b54_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_14":
                                        {
                                            D1b54_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_18":
                                        {
                                            D1b54_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_31":
                                        {
                                            D1b54_31 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_69":
                                        {
                                            D1b54_69 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_71":
                                        {
                                            D1b54_71 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_72":
                                        {
                                            D1b54_72 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_75":
                                        {
                                            D1b54_75 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_76":
                                        {
                                            D1b54_76 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b54_78":
                                        {
                                            D1b54_78 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_1":
                                        {
                                            D1b73_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_4":
                                        {
                                            D1b73_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_6":
                                        {
                                            D1b73_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_14":
                                        {
                                            D1b73_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_18":
                                        {
                                            D1b73_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_31":
                                        {
                                            D1b73_31 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_69":
                                        {
                                            D1b73_69 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_71":
                                        {
                                            D1b73_71 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_72":
                                        {
                                            D1b73_72 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_75":
                                        {
                                            D1b73_75 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_76":
                                        {
                                            D1b73_76 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b73_78":
                                        {
                                            D1b73_78 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_1":
                                        {
                                            D1b75_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_4":
                                        {
                                            D1b75_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_6":
                                        {
                                            D1b75_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_14":
                                        {
                                            D1b75_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_18":
                                        {
                                            D1b75_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_31":
                                        {
                                            D1b75_31 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_69":
                                        {
                                            D1b75_69 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_71":
                                        {
                                            D1b75_71 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_72":
                                        {
                                            D1b75_72 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_75":
                                        {
                                            D1b75_75 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_76":
                                        {
                                            D1b75_76 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b75_78":
                                        {
                                            D1b75_78 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_1":
                                        {
                                            D1b88_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_4":
                                        {
                                            D1b88_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_6":
                                        {
                                            D1b88_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_14":
                                        {
                                            D1b88_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_18":
                                        {
                                            D1b88_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_31":
                                        {
                                            D1b88_31 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_69":
                                        {
                                            D1b88_69 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_71":
                                        {
                                            D1b88_71 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_72":
                                        {
                                            D1b88_72 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_75":
                                        {
                                            D1b88_75 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_76":
                                        {
                                            D1b88_76 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "D1b88_78":
                                        {
                                            D1b88_78 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }


                                }
                            }
                        }
                    }
                    if (UserID != null && Weight != 0)
                    {
                        iobj.insert_Dashboard_values(UserID, SurveyPeriod, MaritalStatus, Country, SurveyID, Gender, Weight, Location, AgeGroup, Edulevel, Occupation, Ses, Period, BevConP6MBotWat, BevConP6MCoffee, BevConP6MFrjuice, BevConP6MIcTea, BevConP6MSoftdrinks, BevConP6MEnDrinks, BevConP6MChocDrink, BevConP6MCerDrink, BevConP6MMilk, BevConP6MHotCoffee, BevConP6MColdCoffee, BevConP6MNoneAb, BevConP3MBotWat, BevConP3MCoffee, BevConP3MFrjuice, BevConP3MIcTea, BevConP3MSoftdrinks, BevConP3MEnDrinks, BevConP3MChocDrink, BevConP3MCerDrink, BevConP3MMilk, BevConP3MHotCoffee, BevConP3MColdCoffee, BevConP3MNoneAb, BevConP1MBotWat, BevConP1MCoffee, BevConP1MFrjuice, BevConP1MIcTea, BevConP1MSoftdrinks, BevConP1MEnDrinks, BevConP1MChocDrink, BevConP1MCerDrink, BevConP1MMilk, BevConP1MHotCoffee, BevConP1MColdCoffee, BevConP1MNoneAb, BevConP1WBotWat, BevConP1WCoffee, BevConP1WFrjuice, BevConP1WIcTea, BevConP1WSoftdrinks, BevConP1WEnDrinks, BevConP1WChocDrink, BevConP1WCerDrink, BevConP1WMilk, BevConP1WHotCoffee, BevConP1WColdCoffee, BevConP1WNoneAb, BrTom, BrSpontKOPIKOBROWN, BrSpontN_CreamyLatte, BrSpontNCW, BrSpontN_Brn_Creamy, BrSpontGTW, BrSpontGTO_TWINPK, BrSpontKOPIKOASTIG, BrSpontKopikoCafBlnca, BrSpontGTW_TwinPk, BrSpontNCW_TwinPk, BrSpontKopikoAstig_TwPk, BrSpontKopikoBrn_TwPk, BrSpontKopikoBlca_TwPk, BrSpontN_Bl_Brew_Br, BrAidKOPIKOBROWN, BrAidN_CreamyLatte, BrAidNCW, BrAidN_Brn_Creamy, BrAidGTW, BrAidGTO_TWINPK, BrAidKOPIKOASTIG, BrAidKopikoCafBlnca, BrAidGTW_TwinPk, BrAidNCW_TwinPk, BrAidKopikoAstig_TwPk, BrAidKopikoBrn_TwPk, BrAidKopikoBlca_TwPk, BrAidN_Bl_Brew_Br, AdTom, AdSpontKOPIKOBROWN, AdSpontN_CreamyLatte, AdSpontNCW, AdSpontN_Brn_Creamy, AdSpontGTW, AdSpontGTO_TWINPK, AdSpontKOPIKOASTIG, AdSpontKopikoCafBlnca, AdSpontGTW_TwinPk, AdSpontNCW_TwinPk, AdSpontKopikoAstig_TwPk, AdSpontKopikoBrn_TwPk, AdSpontKopikoBlca_TwPk, AdSpontN_Bl_Brew_Br, AdAidKOPIKOBROWN, AdAidN_CreamyLatte, AdAidNCW, AdAidN_Brn_Creamy, AdAidGTW, AdAidGTO_TWINPK, AdAidKOPIKOASTIG, AdAidKopikoCafBlnca, AdAidGTW_TwinPk, AdAidNCW_TwinPk, AdAidKopikoAstig_TwPk, AdAidKopikoBrn_TwPk, AdAidKopikoBlca_TwPk, AdAidN_Bl_Brew_Br, CofConP1MRoasted, CofConP1MVend, CofConP1MIntCoffee, CofConP1MIntPowCoff, CofConP1MRTDCoff, CofConP1MCoffCaf, CofConP1MBlackCoff, CofCon1WRoasted, CofCon1WVend, CofCon1WIntCoffee, CofCon1WIntPowCoff, CofCon1WRTDCoff, CofCon1WCoffCaf, CofCon1WBlackCoff, CofCon1DRoasted, CofCon1DVend, CofCon1DIntCoffee, CofCon1DIntPowCoff, CofCon1DRTDCoff, CofCon1DCoffCaf, CofCon1DBlackCoff, EvTriKOPIKOBROWN, EvTriN_CreamyLatte, EvTriNCW, EvTriN_Brn_Creamy, EvTriGTW, EvTriGTO_TWINPK, EvTriKOPIKOASTIG, EvTriKopikoCafBlnca, EvTriGTW_TwinPk, EvTriNCW_TwinPk, EvTriKopikoAstig_TwPk, EvTriKopikoBrn_TwPk, EvTriKopikoBlca_TwPk, EvTriN_Bl_Brew_Br, LaConsKOPIKOBROWN, LaConsN_CreamyLatte, LaConsNCW, LaConsN_Brn_Creamy, LaConsGTW, LaConsGTO_TWINPK, LaConsKOPIKOASTIG, LaConsKopikoCafBlnca, LaConsGTW_TwinPk, LaConsNCW_TwinPk, LaConsKopikoAstig_TwPk, LaConsKopikoBrn_TwPk, LaConsKopikoBlca_TwPk, LaConsN_Bl_Brew_Br, RegP3MKOPIKOBROWN, RegP3MN_CreamyLatte, RegP3MNCW, RegP3MN_Brn_Creamy, RegP3MGTW, RegP3MGTO_TWINPK, RegP3MKOPIKOASTIG, RegP3MKopikoCafBlnca, RegP3MGTW_TwinPk, RegP3MNCW_TwinPk, RegP3MKopikoAstig_TwPk, RegP3MKopikoBrn_TwPk, RegP3MKopikoBlca_TwPk, RegP3MN_Bl_Brew_Br, Bumo, PreBumo, FavBrand, IntSwBr_6M, BrLikeSw, NetBrTom, NetBrSpontGreat_Taste, NetBrSpontKopiko, NetBrSpontNescafe, NetBrAidGreat_Taste, NetBrAidKopiko, NetBrAidNescafe, NetAdTom, NetAdSpontGreat_Taste, NetAdSpontKopiko, NetAdSpontNescafe, NetAdAidGreat_Taste, NetAdAidKopiko, NetAdAidNescafe, NetEvTriGreat_Taste, NetEvTriKopiko, NetEvTriNescafe, NetRegP3MGreat_Taste, NetRegP3MKopiko, NetRegP3MNescafe, NetBumo, NetPreBumo, NetFavBrand, NetBrLikeSw, D1b21_1, D1b21_4, D1b21_6, D1b21_14, D1b21_18, D1b21_31, D1b21_69, D1b21_71, D1b21_72, D1b21_75, D1b21_76, D1b21_78, D1b23_1, D1b23_4, D1b23_6, D1b23_14, D1b23_18, D1b23_31, D1b23_69, D1b23_71, D1b23_72, D1b23_75, D1b23_76, D1b23_78, D1b50_1, D1b50_4, D1b50_6, D1b50_14, D1b50_18, D1b50_31, D1b50_69, D1b50_71, D1b50_72, D1b50_75, D1b50_76, D1b50_78, D1b51_1, D1b51_4, D1b51_6, D1b51_14, D1b51_18, D1b51_31, D1b51_69, D1b51_71, D1b51_72, D1b51_75, D1b51_76, D1b51_78, D1b54_1, D1b54_4, D1b54_6, D1b54_14, D1b54_18, D1b54_31, D1b54_69, D1b54_71, D1b54_72, D1b54_75, D1b54_76, D1b54_78, D1b73_1, D1b73_4, D1b73_6, D1b73_14, D1b73_18, D1b73_31, D1b73_69, D1b73_71, D1b73_72, D1b73_75, D1b73_76, D1b73_78, D1b75_1, D1b75_4, D1b75_6, D1b75_14, D1b75_18, D1b75_31, D1b75_69, D1b75_71, D1b75_72, D1b75_75, D1b75_76, D1b75_78, D1b88_1, D1b88_4, D1b88_6, D1b88_14, D1b88_18, D1b88_31, D1b88_69, D1b88_71, D1b88_72, D1b88_75, D1b88_76, D1b88_78);
                    }
                }

            }
        }

        private static string getYear(string SurveyPeriod)
        {
            string[] date = SurveyPeriod.Split('-');
            return date[0];
        }

        private static string find_UserID(int SurveyID, string SurveyPeriod, string u_id)
        {
            string sum = "";
            string[] date = SurveyPeriod.Split('-');
            foreach (string d in date)
            {
                sum = sum + d;

            }
            return u_id + SurveyID + sum;
        }

    }
    
 
    class Phil_insertion
        {
            SqlConnection connection = new SqlConnection("Data Source=52.74.59.117;Initial Catalog=ClueboxMobile;Persist Security Info=True;User ID=sa;Password=ClueBox123*;");
            internal void insert_Dashboard_variable_values(string VARIABLE_NAME, string VARIABLE_NAME_SUB_NAME, string VARIABLE_ID, string VARIABLE_VALUE, string VARIABLE_NAME_QUESTION, int SurveyID, string country, string BASE_VARIABLE_NAME, string SurveyPeriod)
            {
                connection.Open();
                SqlCommand cmd = new SqlCommand("INSERT INTO DashboardSPS_Variable_Values (VARIABLE_NAME,VARIABLE_NAME_SUB_NAME,VARIABLE_ID,VARIABLE_VALUE,VARIABLE_NAME_QUESTION,SURVEY_ID,SURVEY_COUNTRY,BASE_VARIABLE_NAME,SURVEY_PERIOD) " + "VALUES('" + VARIABLE_NAME + "','" + VARIABLE_NAME_SUB_NAME + "','" + VARIABLE_ID + "','" + VARIABLE_VALUE.Replace("'", "''") + "','" + VARIABLE_NAME_QUESTION + "','" + SurveyID + "','" + country + "','" + BASE_VARIABLE_NAME + "','" + SurveyPeriod + "')", connection);
                try
                {


                    cmd.ExecuteNonQuery();
                    Console.WriteLine("Dashboard variable values inserted successfully");

                    connection.Close();



                }
                catch (Exception)
                {

                    Console.WriteLine("Exception occured");
                    Console.ReadLine();
                }

            }
           internal void insert_Dashboard_values(string UserID, string SurveyPeriod, string MaritalStatus, string Country, int SurveyID, string Gender, decimal Weight, string Location, string AgeGroup, string Edulevel, string Occupation, string Ses, string Period, string BevConP6MBotWat, string BevConP6MCoffee, string BevConP6MFrjuice, string BevConP6MIcTea, string BevConP6MSoftdrinks, string BevConP6MEnDrinks, string BevConP6MChocDrink, string BevConP6MCerDrink, string BevConP6MMilk, string BevConP6MHotCoffee, string BevConP6MColdCoffee, string BevConP6MNoneAb, string BevConP3MBotWat, string BevConP3MCoffee, string BevConP3MFrjuice, string BevConP3MIcTea, string BevConP3MSoftdrinks, string BevConP3MEnDrinks, string BevConP3MChocDrink, string BevConP3MCerDrink, string BevConP3MMilk, string BevConP3MHotCoffee, string BevConP3MColdCoffee, string BevConP3MNoneAb, string BevConP1MBotWat, string BevConP1MCoffee, string BevConP1MFrjuice, string BevConP1MIcTea, string BevConP1MSoftdrinks, string BevConP1MEnDrinks, string BevConP1MChocDrink, string BevConP1MCerDrink, string BevConP1MMilk, string BevConP1MHotCoffee, string BevConP1MColdCoffee, string BevConP1MNoneAb, string BevConP1WBotWat, string BevConP1WCoffee, string BevConP1WFrjuice, string BevConP1WIcTea, string BevConP1WSoftdrinks, string BevConP1WEnDrinks, string BevConP1WChocDrink, string BevConP1WCerDrink, string BevConP1WMilk, string BevConP1WHotCoffee, string BevConP1WColdCoffee, string BevConP1WNoneAb, string BrTom, string BrSpontKOPIKOBROWN, string BrSpontN_CreamyLatte, string BrSpontNCW, string BrSpontN_Brn_Creamy, string BrSpontGTW, string BrSpontGTO_TWINPK, string BrSpontKOPIKOASTIG, string BrSpontKopikoCafBlnca, string BrSpontGTW_TwinPk, string BrSpontNCW_TwinPk, string BrSpontKopikoAstig_TwPk, string BrSpontKopikoBrn_TwPk, string BrSpontKopikoBlca_TwPk, string BrSpontN_Bl_Brew_Br, string BrAidKOPIKOBROWN, string BrAidN_CreamyLatte, string BrAidNCW, string BrAidN_Brn_Creamy, string BrAidGTW, string BrAidGTO_TWINPK, string BrAidKOPIKOASTIG, string BrAidKopikoCafBlnca, string BrAidGTW_TwinPk, string BrAidNCW_TwinPk, string BrAidKopikoAstig_TwPk, string BrAidKopikoBrn_TwPk, string BrAidKopikoBlca_TwPk, string BrAidN_Bl_Brew_Br, string AdTom, string AdSpontKOPIKOBROWN, string AdSpontN_CreamyLatte, string AdSpontNCW, string AdSpontN_Brn_Creamy, string AdSpontGTW, string AdSpontGTO_TWINPK, string AdSpontKOPIKOASTIG, string AdSpontKopikoCafBlnca, string AdSpontGTW_TwinPk, string AdSpontNCW_TwinPk, string AdSpontKopikoAstig_TwPk, string AdSpontKopikoBrn_TwPk, string AdSpontKopikoBlca_TwPk, string AdSpontN_Bl_Brew_Br, string AdAidKOPIKOBROWN, string AdAidN_CreamyLatte, string AdAidNCW, string AdAidN_Brn_Creamy, string AdAidGTW, string AdAidGTO_TWINPK, string AdAidKOPIKOASTIG, string AdAidKopikoCafBlnca, string AdAidGTW_TwinPk, string AdAidNCW_TwinPk, string AdAidKopikoAstig_TwPk, string AdAidKopikoBrn_TwPk, string AdAidKopikoBlca_TwPk, string AdAidN_Bl_Brew_Br, string CofConP1MRoasted, string CofConP1MVend, string CofConP1MIntCoffee, string CofConP1MIntPowCoff, string CofConP1MRTDCoff, string CofConP1MCoffCaf, string CofConP1MBlackCoff, string CofCon1WRoasted, string CofCon1WVend, string CofCon1WIntCoffee, string CofCon1WIntPowCoff, string CofCon1WRTDCoff, string CofCon1WCoffCaf, string CofCon1WBlackCoff, string CofCon1DRoasted, string CofCon1DVend, string CofCon1DIntCoffee, string CofCon1DIntPowCoff, string CofCon1DRTDCoff, string CofCon1DCoffCaf, string CofCon1DBlackCoff, string EvTriKOPIKOBROWN, string EvTriN_CreamyLatte, string EvTriNCW, string EvTriN_Brn_Creamy, string EvTriGTW, string EvTriGTO_TWINPK, string EvTriKOPIKOASTIG, string EvTriKopikoCafBlnca, string EvTriGTW_TwinPk, string EvTriNCW_TwinPk, string EvTriKopikoAstig_TwPk, string EvTriKopikoBrn_TwPk, string EvTriKopikoBlca_TwPk, string EvTriN_Bl_Brew_Br, string LaConsKOPIKOBROWN, string LaConsN_CreamyLatte, string LaConsNCW, string LaConsN_Brn_Creamy, string LaConsGTW, string LaConsGTO_TWINPK, string LaConsKOPIKOASTIG, string LaConsKopikoCafBlnca, string LaConsGTW_TwinPk, string LaConsNCW_TwinPk, string LaConsKopikoAstig_TwPk, string LaConsKopikoBrn_TwPk, string LaConsKopikoBlca_TwPk, string LaConsN_Bl_Brew_Br, string RegP3MKOPIKOBROWN, string RegP3MN_CreamyLatte, string RegP3MNCW, string RegP3MN_Brn_Creamy, string RegP3MGTW, string RegP3MGTO_TWINPK, string RegP3MKOPIKOASTIG, string RegP3MKopikoCafBlnca, string RegP3MGTW_TwinPk, string RegP3MNCW_TwinPk, string RegP3MKopikoAstig_TwPk, string RegP3MKopikoBrn_TwPk, string RegP3MKopikoBlca_TwPk, string RegP3MN_Bl_Brew_Br, string Bumo, string PreBumo, string FavBrand, string IntSwBr_6M, string BrLikeSw, string NetBrTom, string NetBrSpontGreat_Taste, string NetBrSpontKopiko, string NetBrSpontNescafe, string NetBrAidGreat_Taste, string NetBrAidKopiko, string NetBrAidNescafe, string NetAdTom, string NetAdSpontGreat_Taste, string NetAdSpontKopiko, string NetAdSpontNescafe, string NetAdAidGreat_Taste, string NetAdAidKopiko, string NetAdAidNescafe, string NetEvTriGreat_Taste, string NetEvTriKopiko, string NetEvTriNescafe, string NetRegP3MGreat_Taste, string NetRegP3MKopiko, string NetRegP3MNescafe, string NetBumo, string NetPreBumo, string NetFavBrand, string NetBrLikeSw, string D1b21_1, string D1b21_4, string D1b21_6, string D1b21_14, string D1b21_18, string D1b21_31, string D1b21_69, string D1b21_71, string D1b21_72, string D1b21_75, string D1b21_76, string D1b21_78, string D1b23_1, string D1b23_4, string D1b23_6, string D1b23_14, string D1b23_18, string D1b23_31, string D1b23_69, string D1b23_71, string D1b23_72, string D1b23_75, string D1b23_76, string D1b23_78, string D1b50_1, string D1b50_4, string D1b50_6, string D1b50_14, string D1b50_18, string D1b50_31, string D1b50_69, string D1b50_71, string D1b50_72, string D1b50_75, string D1b50_76, string D1b50_78, string D1b51_1, string D1b51_4, string D1b51_6, string D1b51_14, string D1b51_18, string D1b51_31, string D1b51_69, string D1b51_71, string D1b51_72, string D1b51_75, string D1b51_76, string D1b51_78, string D1b54_1, string D1b54_4, string D1b54_6, string D1b54_14, string D1b54_18, string D1b54_31, string D1b54_69, string D1b54_71, string D1b54_72, string D1b54_75, string D1b54_76, string D1b54_78, string D1b73_1, string D1b73_4, string D1b73_6, string D1b73_14, string D1b73_18, string D1b73_31, string D1b73_69, string D1b73_71, string D1b73_72, string D1b73_75, string D1b73_76, string D1b73_78, string D1b75_1, string D1b75_4, string D1b75_6, string D1b75_14, string D1b75_18, string D1b75_31, string D1b75_69, string D1b75_71, string D1b75_72, string D1b75_75, string D1b75_76, string D1b75_78, string D1b88_1, string D1b88_4, string D1b88_6, string D1b88_14, string D1b88_18, string D1b88_31, string D1b88_69, string D1b88_71, string D1b88_72, string D1b88_75, string D1b88_76, string D1b88_78)
            {
                int i;
                connection.Open();
                //SqlConnection connection = new SqlConnection("Data Source=52.74.59.117;Initial Catalog=ClueboxMobile;Persist Security Info=True;User ID=sa;Password=ClueBox123*;");
                SqlCommand cmd = new SqlCommand("INSERT INTO DashboardFlatTab_Phil_COFF (UserID,AttendedOn,MaritalStatus,Country,SurveyID,Gender,Weight,Location,AgeGroup,Edulevel,Occupation,Ses,Period,BevConP6MBotWat,BevConP6MCoffee,BevConP6MFrjuice,BevConP6MIcTea,BevConP6MSoftdrinks,BevConP6MEnDrinks,BevConP6MChocDrink,BevConP6MCerDrink,BevConP6MMilk,BevConP6MHotCoffee,BevConP6MColdCoffee,BevConP6MNoneAb,BevConP3MBotWat,BevConP3MCoffee,BevConP3MFrjuice,BevConP3MIcTea,BevConP3MSoftdrinks,BevConP3MEnDrinks,BevConP3MChocDrink,BevConP3MCerDrink,BevConP3MMilk,BevConP3MHotCoffee,BevConP3MColdCoffee,BevConP3MNoneAb,BevConP1MBotWat,BevConP1MCoffee,BevConP1MFrjuice,BevConP1MIcTea,BevConP1MSoftdrinks,BevConP1MEnDrinks,BevConP1MChocDrink,BevConP1MCerDrink,BevConP1MMilk,BevConP1MHotCoffee,BevConP1MColdCoffee,BevConP1MNoneAb,BevConP1WBotWat,BevConP1WCoffee,BevConP1WFrjuice,BevConP1WIcTea,BevConP1WSoftdrinks,BevConP1WEnDrinks,BevConP1WChocDrink,BevConP1WCerDrink,BevConP1WMilk,BevConP1WHotCoffee,BevConP1WColdCoffee,BevConP1WNoneAb,BrTom,BrSpontKOPIKOBROWN,BrSpontN_CreamyLatte,BrSpontNCW,BrSpontN_Brn_Creamy,BrSpontGTW,BrSpontGTO_TWINPK,BrSpontKOPIKOASTIG,BrSpontKopikoCafBlnca,BrSpontGTW_TwinPk,BrSpontNCW_TwinPk,BrSpontKopikoAstig_TwPk,BrSpontKopikoBrn_TwPk,BrSpontKopikoBlca_TwPk,BrSpontN_Bl_Brew_Br,BrAidKOPIKOBROWN,BrAidN_CreamyLatte,BrAidNCW,BrAidN_Brn_Creamy,BrAidGTW,BrAidGTO_TWINPK,BrAidKOPIKOASTIG,BrAidKopikoCafBlnca,BrAidGTW_TwinPk,BrAidNCW_TwinPk,BrAidKopikoAstig_TwPk,BrAidKopikoBrn_TwPk,BrAidKopikoBlca_TwPk,BrAidN_Bl_Brew_Br,AdTom,AdSpontKOPIKOBROWN,AdSpontN_CreamyLatte,AdSpontNCW,AdSpontN_Brn_Creamy,AdSpontGTW,AdSpontGTO_TWINPK,AdSpontKOPIKOASTIG,AdSpontKopikoCafBlnca,AdSpontGTW_TwinPk,AdSpontNCW_TwinPk,AdSpontKopikoAstig_TwPk,AdSpontKopikoBrn_TwPk,AdSpontKopikoBlca_TwPk,AdSpontN_Bl_Brew_Br,AdAidKOPIKOBROWN,AdAidN_CreamyLatte,AdAidNCW,AdAidN_Brn_Creamy,AdAidGTW,AdAidGTO_TWINPK,AdAidKOPIKOASTIG,AdAidKopikoCafBlnca,AdAidGTW_TwinPk,AdAidNCW_TwinPk,AdAidKopikoAstig_TwPk,AdAidKopikoBrn_TwPk,AdAidKopikoBlca_TwPk,AdAidN_Bl_Brew_Br,CofConP1MRoasted,CofConP1MVend,CofConP1MIntCoffee,CofConP1MIntPowCoff,CofConP1MRTDCoff,CofConP1MCoffCaf,CofConP1MBlackCoff,CofCon1WRoasted,CofCon1WVend,CofCon1WIntCoffee,CofCon1WIntPowCoff,CofCon1WRTDCoff,CofCon1WCoffCaf,CofCon1WBlackCoff,CofCon1DRoasted,CofCon1DVend,CofCon1DIntCoffee,CofCon1DIntPowCoff,CofCon1DRTDCoff,CofCon1DCoffCaf,CofCon1DBlackCoff,EvTriKOPIKOBROWN,EvTriN_CreamyLatte,EvTriNCW,EvTriN_Brn_Creamy,EvTriGTW,EvTriGTO_TWINPK,EvTriKOPIKOASTIG,EvTriKopikoCafBlnca,EvTriGTW_TwinPk,EvTriNCW_TwinPk,EvTriKopikoAstig_TwPk,EvTriKopikoBrn_TwPk,EvTriKopikoBlca_TwPk,EvTriN_Bl_Brew_Br,LaConsKOPIKOBROWN,LaConsN_CreamyLatte,LaConsNCW,LaConsN_Brn_Creamy,LaConsGTW,LaConsGTO_TWINPK,LaConsKOPIKOASTIG,LaConsKopikoCafBlnca,LaConsGTW_TwinPk,LaConsNCW_TwinPk,LaConsKopikoAstig_TwPk,LaConsKopikoBrn_TwPk,LaConsKopikoBlca_TwPk,LaConsN_Bl_Brew_Br,RegP3MKOPIKOBROWN,RegP3MN_CreamyLatte,RegP3MNCW,RegP3MN_Brn_Creamy,RegP3MGTW,RegP3MGTO_TWINPK,RegP3MKOPIKOASTIG,RegP3MKopikoCafBlnca,RegP3MGTW_TwinPk,RegP3MNCW_TwinPk,RegP3MKopikoAstig_TwPk,RegP3MKopikoBrn_TwPk,RegP3MKopikoBlca_TwPk,RegP3MN_Bl_Brew_Br,Bumo,PreBumo,FavBrand,IntSwBr_6M,BrLikeSw,NetBrTom,NetBrSpontGreat_Taste,NetBrSpontKopiko,NetBrSpontNescafe,NetBrAidGreat_Taste,NetBrAidKopiko,NetBrAidNescafe,NetAdTom,NetAdSpontGreat_Taste,NetAdSpontKopiko,NetAdSpontNescafe,NetAdAidGreat_Taste,NetAdAidKopiko,NetAdAidNescafe,NetEvTriGreat_Taste,NetEvTriKopiko,NetEvTriNescafe,NetRegP3MGreat_Taste,NetRegP3MKopiko,NetRegP3MNescafe,NetBumo,NetPreBumo,NetFavBrand,NetBrLikeSw,D1b21_1,D1b21_4,D1b21_6,D1b21_14,D1b21_18,D1b21_31,D1b21_69,D1b21_71,D1b21_72,D1b21_75,D1b21_76,D1b21_78,D1b23_1,D1b23_4,D1b23_6,D1b23_14,D1b23_18,D1b23_31,D1b23_69,D1b23_71,D1b23_72,D1b23_75,D1b23_76,D1b23_78,D1b50_1,D1b50_4,D1b50_6,D1b50_14,D1b50_18,D1b50_31,D1b50_69,D1b50_71,D1b50_72,D1b50_75,D1b50_76,D1b50_78,D1b51_1,D1b51_4,D1b51_6,D1b51_14,D1b51_18,D1b51_31,D1b51_69,D1b51_71,D1b51_72,D1b51_75,D1b51_76,D1b51_78,D1b54_1,D1b54_4,D1b54_6,D1b54_14,D1b54_18,D1b54_31,D1b54_69,D1b54_71,D1b54_72,D1b54_75,D1b54_76,D1b54_78,D1b73_1,D1b73_4,D1b73_6,D1b73_14,D1b73_18,D1b73_31,D1b73_69,D1b73_71,D1b73_72,D1b73_75,D1b73_76,D1b73_78,D1b75_1,D1b75_4,D1b75_6,D1b75_14,D1b75_18,D1b75_31,D1b75_69,D1b75_71,D1b75_72,D1b75_75,D1b75_76,D1b75_78,D1b88_1,D1b88_4,D1b88_6,D1b88_14,D1b88_18,D1b88_31,D1b88_69,D1b88_71,D1b88_72,D1b88_75,D1b88_76,D1b88_78) " + "VALUES('" + UserID + "','" + SurveyPeriod + "','" + MaritalStatus + "','" + Country + "','" + SurveyID + "','" + Gender + "','" + Weight + "','" + Location + "','" + AgeGroup + "','" + Edulevel + "','" + Occupation + "','" + Ses + "','" + Period + "','" + BevConP6MBotWat + "','" + BevConP6MCoffee + "','" + BevConP6MFrjuice + "','" + BevConP6MIcTea + "','" + BevConP6MSoftdrinks + "','" + BevConP6MEnDrinks + "','" + BevConP6MChocDrink + "','" + BevConP6MCerDrink + "','" + BevConP6MMilk + "','" + BevConP6MHotCoffee + "','" + BevConP6MColdCoffee + "','" + BevConP6MNoneAb + "','" + BevConP3MBotWat + "','" + BevConP3MCoffee + "','" + BevConP3MFrjuice + "','" + BevConP3MIcTea + "','" + BevConP3MSoftdrinks + "','" + BevConP3MEnDrinks + "','" + BevConP3MChocDrink + "','" + BevConP3MCerDrink + "','" + BevConP3MMilk + "','" + BevConP3MHotCoffee + "','" + BevConP3MColdCoffee + "','" + BevConP3MNoneAb + "','" + BevConP1MBotWat + "','" + BevConP1MCoffee + "','" + BevConP1MFrjuice + "','" + BevConP1MIcTea + "','" + BevConP1MSoftdrinks + "','" + BevConP1MEnDrinks + "','" + BevConP1MChocDrink + "','" + BevConP1MCerDrink + "','" + BevConP1MMilk + "','" + BevConP1MHotCoffee + "','" + BevConP1MColdCoffee + "','" + BevConP1MNoneAb + "','" + BevConP1WBotWat + "','" + BevConP1WCoffee + "','" + BevConP1WFrjuice + "','" + BevConP1WIcTea + "','" + BevConP1WSoftdrinks + "','" + BevConP1WEnDrinks + "','" + BevConP1WChocDrink + "','" + BevConP1WCerDrink + "','" + BevConP1WMilk + "','" + BevConP1WHotCoffee + "','" + BevConP1WColdCoffee + "','" + BevConP1WNoneAb + "','" + BrTom + "','" + BrSpontKOPIKOBROWN + "','" + BrSpontN_CreamyLatte + "','" + BrSpontNCW + "','" + BrSpontN_Brn_Creamy + "','" + BrSpontGTW + "','" + BrSpontGTO_TWINPK + "','" + BrSpontKOPIKOASTIG + "','" + BrSpontKopikoCafBlnca + "','" + BrSpontGTW_TwinPk + "','" + BrSpontNCW_TwinPk + "','" + BrSpontKopikoAstig_TwPk + "','" + BrSpontKopikoBrn_TwPk + "','" + BrSpontKopikoBlca_TwPk + "','" + BrSpontN_Bl_Brew_Br + "','" + BrAidKOPIKOBROWN + "','" + BrAidN_CreamyLatte + "','" + BrAidNCW + "','" + BrAidN_Brn_Creamy + "','" + BrAidGTW + "','" + BrAidGTO_TWINPK + "','" + BrAidKOPIKOASTIG + "','" + BrAidKopikoCafBlnca + "','" + BrAidGTW_TwinPk + "','" + BrAidNCW_TwinPk + "','" + BrAidKopikoAstig_TwPk + "','" + BrAidKopikoBrn_TwPk + "','" + BrAidKopikoBlca_TwPk + "','" + BrAidN_Bl_Brew_Br + "','" + AdTom + "','" + AdSpontKOPIKOBROWN + "','" + AdSpontN_CreamyLatte + "','" + AdSpontNCW + "','" + AdSpontN_Brn_Creamy + "','" + AdSpontGTW + "','" + AdSpontGTO_TWINPK + "','" + AdSpontKOPIKOASTIG + "','" + AdSpontKopikoCafBlnca + "','" + AdSpontGTW_TwinPk + "','" + AdSpontNCW_TwinPk + "','" + AdSpontKopikoAstig_TwPk + "','" + AdSpontKopikoBrn_TwPk + "','" + AdSpontKopikoBlca_TwPk + "','" + AdSpontN_Bl_Brew_Br + "','" + AdAidKOPIKOBROWN + "','" + AdAidN_CreamyLatte + "','" + AdAidNCW + "','" + AdAidN_Brn_Creamy + "','" + AdAidGTW + "','" + AdAidGTO_TWINPK + "','" + AdAidKOPIKOASTIG + "','" + AdAidKopikoCafBlnca + "','" + AdAidGTW_TwinPk + "','" + AdAidNCW_TwinPk + "','" + AdAidKopikoAstig_TwPk + "','" + AdAidKopikoBrn_TwPk + "','" + AdAidKopikoBlca_TwPk + "','" + AdAidN_Bl_Brew_Br + "','" + CofConP1MRoasted + "','" + CofConP1MVend + "','" + CofConP1MIntCoffee + "','" + CofConP1MIntPowCoff + "','" + CofConP1MRTDCoff + "','" + CofConP1MCoffCaf + "','" + CofConP1MBlackCoff + "','" + CofCon1WRoasted + "','" + CofCon1WVend + "','" + CofCon1WIntCoffee + "','" + CofCon1WIntPowCoff + "','" + CofCon1WRTDCoff + "','" + CofCon1WCoffCaf + "','" + CofCon1WBlackCoff + "','" + CofCon1DRoasted + "','" + CofCon1DVend + "','" + CofCon1DIntCoffee + "','" + CofCon1DIntPowCoff + "','" + CofCon1DRTDCoff + "','" + CofCon1DCoffCaf + "','" + CofCon1DBlackCoff + "','" + EvTriKOPIKOBROWN + "','" + EvTriN_CreamyLatte + "','" + EvTriNCW + "','" + EvTriN_Brn_Creamy + "','" + EvTriGTW + "','" + EvTriGTO_TWINPK + "','" + EvTriKOPIKOASTIG + "','" + EvTriKopikoCafBlnca + "','" + EvTriGTW_TwinPk + "','" + EvTriNCW_TwinPk + "','" + EvTriKopikoAstig_TwPk + "','" + EvTriKopikoBrn_TwPk + "','" + EvTriKopikoBlca_TwPk + "','" + EvTriN_Bl_Brew_Br + "','" + LaConsKOPIKOBROWN + "','" + LaConsN_CreamyLatte + "','" + LaConsNCW + "','" + LaConsN_Brn_Creamy + "','" + LaConsGTW + "','" + LaConsGTO_TWINPK + "','" + LaConsKOPIKOASTIG + "','" + LaConsKopikoCafBlnca + "','" + LaConsGTW_TwinPk + "','" + LaConsNCW_TwinPk + "','" + LaConsKopikoAstig_TwPk + "','" + LaConsKopikoBrn_TwPk + "','" + LaConsKopikoBlca_TwPk + "','" + LaConsN_Bl_Brew_Br + "','" + RegP3MKOPIKOBROWN + "','" + RegP3MN_CreamyLatte + "','" + RegP3MNCW + "','" + RegP3MN_Brn_Creamy + "','" + RegP3MGTW + "','" + RegP3MGTO_TWINPK + "','" + RegP3MKOPIKOASTIG + "','" + RegP3MKopikoCafBlnca + "','" + RegP3MGTW_TwinPk + "','" + RegP3MNCW_TwinPk + "','" + RegP3MKopikoAstig_TwPk + "','" + RegP3MKopikoBrn_TwPk + "','" + RegP3MKopikoBlca_TwPk + "','" + RegP3MN_Bl_Brew_Br + "','" + Bumo + "','" + PreBumo + "','" + FavBrand + "','" + IntSwBr_6M + "','" + BrLikeSw + "','" + NetBrTom + "','" + NetBrSpontGreat_Taste + "','" + NetBrSpontKopiko + "','" + NetBrSpontNescafe + "','" + NetBrAidGreat_Taste + "','" + NetBrAidKopiko + "','" + NetBrAidNescafe + "','" + NetAdTom + "','" + NetAdSpontGreat_Taste + "','" + NetAdSpontKopiko + "','" + NetAdSpontNescafe + "','" + NetAdAidGreat_Taste + "','" + NetAdAidKopiko + "','" + NetAdAidNescafe + "','" + NetEvTriGreat_Taste + "','" + NetEvTriKopiko + "','" + NetEvTriNescafe + "','" + NetRegP3MGreat_Taste + "','" + NetRegP3MKopiko + "','" + NetRegP3MNescafe + "','" + NetBumo + "','" + NetPreBumo + "','" + NetFavBrand + "','" + NetBrLikeSw + "','" + D1b21_1 + "','" + D1b21_4 + "','" + D1b21_6 + "','" + D1b21_14 + "','" + D1b21_18 + "','" + D1b21_31 + "','" + D1b21_69 + "','" + D1b21_71 + "','" + D1b21_72 + "','" + D1b21_75 + "','" + D1b21_76 + "','" + D1b21_78 + "','" + D1b23_1 + "','" + D1b23_4 + "','" + D1b23_6 + "','" + D1b23_14 + "','" + D1b23_18 + "','" + D1b23_31 + "','" + D1b23_69 + "','" + D1b23_71 + "','" + D1b23_72 + "','" + D1b23_75 + "','" + D1b23_76 + "','" + D1b23_78 + "','" + D1b50_1 + "','" + D1b50_4 + "','" + D1b50_6 + "','" + D1b50_14 + "','" + D1b50_18 + "','" + D1b50_31 + "','" + D1b50_69 + "','" + D1b50_71 + "','" + D1b50_72 + "','" + D1b50_75 + "','" + D1b50_76 + "','" + D1b50_78 + "','" + D1b51_1 + "','" + D1b51_4 + "','" + D1b51_6 + "','" + D1b51_14 + "','" + D1b51_18 + "','" + D1b51_31 + "','" + D1b51_69 + "','" + D1b51_71 + "','" + D1b51_72 + "','" + D1b51_75 + "','" + D1b51_76 + "','" + D1b51_78 + "','" + D1b54_1 + "','" + D1b54_4 + "','" + D1b54_6 + "','" + D1b54_14 + "','" + D1b54_18 + "','" + D1b54_31 + "','" + D1b54_69 + "','" + D1b54_71 + "','" + D1b54_72 + "','" + D1b54_75 + "','" + D1b54_76 + "','" + D1b54_78 + "','" + D1b73_1 + "','" + D1b73_4 + "','" + D1b73_6 + "','" + D1b73_14 + "','" + D1b73_18 + "','" + D1b73_31 + "','" + D1b73_69 + "','" + D1b73_71 + "','" + D1b73_72 + "','" + D1b73_75 + "','" + D1b73_76 + "','" + D1b73_78 + "','" + D1b75_1 + "','" + D1b75_4 + "','" + D1b75_6 + "','" + D1b75_14 + "','" + D1b75_18 + "','" + D1b75_31 + "','" + D1b75_69 + "','" + D1b75_71 + "','" + D1b75_72 + "','" + D1b75_75 + "','" + D1b75_76 + "','" + D1b75_78 + "','" + D1b88_1 + "','" + D1b88_4 + "','" + D1b88_6 + "','" + D1b88_14 + "','" + D1b88_18 + "','" + D1b88_31 + "','" + D1b88_69 + "','" + D1b88_71 + "','" + D1b88_72 + "','" + D1b88_75 + "','" + D1b88_76 + "','" + D1b88_78 + "' )", connection);
                try
                {
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("Data inserted successfully");
                    i = 1;
                }
                catch (Exception ex)
                {

                    connection.Close();
                    i = 0;
                    Console.WriteLine("Exception occured" + ex.ToString());
                    Console.WriteLine("Exception occured");

                    Console.ReadLine();
                }
                connection.Close();
            }
      }
 
}

