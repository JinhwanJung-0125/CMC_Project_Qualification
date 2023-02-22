using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Collections.Generic;
using NPOI.SS.Formula.Functions;

/*
 23.01.31 업데이트
 ------------------
 새로운 Xml 구조에 맞게 수정
 ==================
 T3
     C24 -> C9
     C4  -> C5
     C15 -> C16
     C16 -> C17
     C17 -> C18
     C18 -> C19
     C19 -> C20
     C20 -> C21
     C21 -> C22
     C22 -> C23
 ==================
 T5
     C9  -> C4
     C22 -> C8
 ==================
*/
/*
 23.01.31 업데이트2
 --------------------
  사업자등록번호 <T1></C17><T1>에 추가
  SetBusinessInfo() 추가
  Calculation() 내부에서 SetBusinessInfo() 호출 추가
 --------------------
 */
/*
 23.02.02 업데이트
 --------------------
 작업 폴더 경로 수정
 --------------------
*/
/*
 23.02.02 업데이트2
 --------------------
 기존 폴더가 존재해도 제대로 작동되도록 수정
 --------------------
*/
/*
 23.02.06 업데이트
 --------------------
 고정금액 소수점 5자리에서 절사되도록 수정
 --------------------
*/
/*
 23.02.07 업데이트
 --------------------
 공종 합계 저장 메소드 (SetPriceOfSuperConstruction) 추가
 --------------------
*/
/*
 --------------------
 노무비 80% 미만시 단가조정 메소드 추가 (CheckLaborLimit80)
 --------------------
 */


namespace SetUnitPriceByExcel
{
    class CalculatePrice
    {
        static XDocument docBID;
        static IEnumerable<XElement> eleBID;

        static void CalculateResultPrice()  //입찰 금액 계산
        {
            if (Data.ResultPrice == 0)  //입찰 금액을 입력받지 않았다면 금액을 계산한다.
            {
                decimal expectedPrice = (decimal)Data.BasePrice * (Data.PersonalRate + 1);  //예정 가격 계산
                Data.ResultPrice = 0.88m * (expectedPrice - (decimal)Data.priceA) + (decimal)Data.priceA;   //입찰 금액 계산
            }
        }

        static void CalculateEvaluationPrice()  //자제인력 평가 항목 금액 계산
        {
            decimal TotalLabor = Data.ResultPrice * Convert.ToDecimal(Data.laborPercent / 100) * 1.01m; //노무비 기준율 대비 101% 수준의 노무비 금액을 계산한다.
            decimal TotalExpence = Data.ResultPrice * Convert.ToDecimal(Data.expencePercent / 100) * 0.81m; //기타 경비 기준율 대비 81% 수준의 기타 경비 금액을 계산한다.
            decimal TotalManagement = Data.ResultPrice * Convert.ToDecimal(Data.managementPercent / 100) * 0.81m;   //일반 관리비 기준율 대비 81% 수준의 일반 관리비 금액을 계산한다.
            decimal TotalProfit = Data.ResultPrice * Convert.ToDecimal(Data.profitPercent / 100) * 0.81m;   //이윤 기준율 대비 81% 수준의 이윤을 계산한다.

            decimal DirectLabor = TotalLabor / (1 + 0.113m);    //구한 노무비 기준의 직접 노무비를 계산한다.

            Data.Rate = (DirectLabor - Data.FixedPriceDirectLabor + Data.StandardLabor) / (Data.RealPriceDirectLabor + Data.StandardLabor); //직접 노무비 기준의 네고율을 구한다.

            Data.Bidding["기타경비"] = FillCostAccount.ToLong(TotalExpence);    //계산된 기타 경비를 Data에 저장한다.
            Data.Bidding["일반관리비"] = FillCostAccount.ToLong(TotalManagement);    //계산된 일반 관리비를 Data에 저장한다.
            Data.Bidding["이윤"] = FillCostAccount.ToLong(TotalProfit);   //계산된 이윤을 Data에 저장한다.
        }

        static void RoundOrTruncate(decimal Rate, Data Object, ref decimal myMaterialUnit, ref decimal myLaborUnit, ref decimal myExpenseUnit)
        { 
            //단가는 모두 정수로 처리한다.
            myMaterialUnit = Math.Ceiling(Object.MaterialUnit * Rate);
            myLaborUnit = Math.Ceiling(Object.LaborUnit * Rate);
            myExpenseUnit = Math.Ceiling(Object.ExpenseUnit * Rate);
        }
        
        static void Recalculation() //사정율에 따라 재계산된 가격을 비드파일에 복사
        {
            foreach (var bid in eleBID)
            {
                //일반 항목인 경우
                if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")
                {
                    var constNum = string.Concat(bid.Element("C1").Value);      //세부공사 번호
                    var numVal = string.Concat(bid.Element("C2").Value);        //세부공종 번호
                    var detailVal = string.Concat(bid.Element("C3").Value);     //세부 공종 번호
                    var curObject = Data.Dic[constNum].Find(x => x.WorkNum == numVal && x.DetailWorkNum == detailVal);
                    if (curObject.Item.Equals("일반") || curObject.Item.Equals("표준시장단가"))
                    {
                        //직접공사비 재계산
                        Data.RealDirectMaterial -= Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.RealDirectLabor -= Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.RealOutputExpense -= Convert.ToDecimal(string.Concat(bid.Element("C22").Value)); 

                        decimal myMaterialUnit = 0;
                        decimal myLaborUnit = 0;
                        decimal myExpenseUnit = 0;

                        RoundOrTruncate(Data.Rate, curObject, ref myMaterialUnit, ref myLaborUnit, ref myExpenseUnit);  //계산된 네고율로 단가 처리한다.

                        curObject.MaterialUnit = myMaterialUnit;
                        curObject.LaborUnit = myLaborUnit;
                        curObject.ExpenseUnit = myExpenseUnit;

                        //최종 단가 및 합계 계산
                        bid.Element("C16").Value = curObject.MaterialUnit.ToString();    //재료비 단가
                        bid.Element("C17").Value = curObject.LaborUnit.ToString();       //노무비 단가
                        bid.Element("C18").Value = curObject.ExpenseUnit.ToString();     //경비 단가
                        bid.Element("C19").Value = curObject.UnitPriceSum.ToString();    //합계 단가
                        bid.Element("C20").Value = curObject.Material.ToString();    //재료비
                        bid.Element("C21").Value = curObject.Labor.ToString();       //노무비
                        bid.Element("C22").Value = curObject.Expense.ToString();     //경비
                        bid.Element("C23").Value = curObject.PriceSum.ToString();    //합계

                        //붙여넣기한 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                        Data.RealDirectMaterial += Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.RealDirectLabor += Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.RealOutputExpense += Convert.ToDecimal(string.Concat(bid.Element("C22").Value));
                    }
                    if (curObject.Item.Equals("제요율적용제외"))   //제요율적용제외 공종도 네고가 가능하다.
                    {
                        decimal myMaterialUnit = 0;
                        decimal myLaborUnit = 0;
                        decimal myExpenseUnit = 0;

                        RoundOrTruncate(Data.Rate, curObject, ref myMaterialUnit, ref myLaborUnit, ref myExpenseUnit);

                        curObject.MaterialUnit = myMaterialUnit;
                        curObject.LaborUnit = myLaborUnit;
                        curObject.ExpenseUnit = myExpenseUnit;

                        //최종 단가 및 합계 계산
                        bid.Element("C16").Value = curObject.MaterialUnit.ToString();    //재료비 단가
                        bid.Element("C17").Value = curObject.LaborUnit.ToString();       //노무비 단가
                        bid.Element("C18").Value = curObject.ExpenseUnit.ToString();     //경비 단가
                        bid.Element("C19").Value = curObject.UnitPriceSum.ToString();    //합계 단가
                        bid.Element("C20").Value = curObject.Material.ToString();    //재료비
                        bid.Element("C21").Value = curObject.Labor.ToString();       //노무비
                        bid.Element("C22").Value = curObject.Expense.ToString();     //경비
                        bid.Element("C23").Value = curObject.PriceSum.ToString();    //합계

                        Data.AdjustedExMaterial += Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.AdjustedExLabor += Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.AdjustedExExpense += Convert.ToDecimal(string.Concat(bid.Element("C22").Value));
                    }
                }
            }
        }
        

        public static void SetBusinessInfo()
        {
            
            foreach (var bid in eleBID)
            {
                if (bid.Name == "T1")
                {
                    bid.Element("C17").Value = Data.CompanyRegistrationNum;
                    bid.Element("C18").Value = Data.CompanyRegistrationName;
                }
            }

        }

        public static void SetPriceOfSuperConstruction()    //상위 공종의 각 단가 합 및 합계 세팅 (23.02.07)
        {
            XElement? firstConstruction = null;     //가장 상위 공종
            XElement? secondConstruction = null;    //중간 상위 공종
            XElement? thirdConstruction = null;     //마지막 상위 공종

            foreach (var bid in eleBID)
            {
                if(bid.Name == "T3")
                {
                    if (string.Concat(bid.Element("C5").Value) == "G")  //공종이면
                    {
                        if (bid.Element("C23").Value == "0")    //이미 합계가 세팅되어 있는지 확인 (중복 계산을 막기 위함)
                        {
                            if (firstConstruction == null || string.Concat(bid.Element("C3").Value) == "0") //C3이 0이면 가장 상위 공종
                            {
                                firstConstruction = bid;    //현재 보고있는 object가 가장 상위 공종
                                secondConstruction = null;  //중간 상위 공종 초기화
                                thirdConstruction = null;   //마지막 상위 공종 초기화
                            }
                            else if (string.Concat(bid.Element("C3").Value) == string.Concat(firstConstruction.Element("C2").Value) && firstConstruction != null)   //C3이 가장 상위 공종의 C2와 같다면 중간 상위 공종
                            {
                                secondConstruction = bid;   //현재 보고있는 object가 중간 상위 공종
                                thirdConstruction = null;   //마지막 상위 공종 초기화
                            }
                            else if (string.Concat(bid.Element("C3").Value) == string.Concat(secondConstruction.Element("C2").Value) && secondConstruction != null) // C3이 중간 상위 공종의 C2와 같다면 마지막 상위 공종
                                thirdConstruction = bid;    //현재 보고있는 object가 마지막 상위 공종
                        }
                        else   //공종에 합계가 이미 세팅되어 있다면 전부 초기화
                        {
                            firstConstruction = null;
                            secondConstruction = null;
                            thirdConstruction = null;
                        }
                    }
                    else if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")    //공종이 아니면
                    {
                        if (firstConstruction != null)  //현재 보는 object가 가장 상위 공종에 포함되어 있다면 단가별 합과 합계를 더해나감
                        {
                            firstConstruction.Element("C20").Value = string.Concat(Convert.ToDecimal(firstConstruction.Element("C20").Value) + Convert.ToDecimal(bid.Element("C20").Value));    //재료비
                            firstConstruction.Element("C21").Value = string.Concat(Convert.ToDecimal(firstConstruction.Element("C21").Value) + Convert.ToDecimal(bid.Element("C21").Value));    //노무비
                            firstConstruction.Element("C22").Value = string.Concat(Convert.ToDecimal(firstConstruction.Element("C22").Value) + Convert.ToDecimal(bid.Element("C22").Value));    //경비
                            firstConstruction.Element("C23").Value = string.Concat(Convert.ToDecimal(firstConstruction.Element("C23").Value) + Convert.ToDecimal(bid.Element("C23").Value));    //합계
                        }
                        if (secondConstruction != null) //현재 보는 object가 중간 상위 공종에 포함되어 있다면 단가별 합과 합계를 더해나감
                        {
                            secondConstruction.Element("C20").Value = string.Concat(Convert.ToDecimal(secondConstruction.Element("C20").Value) + Convert.ToDecimal(bid.Element("C20").Value));  //재료비
                            secondConstruction.Element("C21").Value = string.Concat(Convert.ToDecimal(secondConstruction.Element("C21").Value) + Convert.ToDecimal(bid.Element("C21").Value));  //노무비
                            secondConstruction.Element("C22").Value = string.Concat(Convert.ToDecimal(secondConstruction.Element("C22").Value) + Convert.ToDecimal(bid.Element("C22").Value));  //경비
                            secondConstruction.Element("C23").Value = string.Concat(Convert.ToDecimal(secondConstruction.Element("C23").Value) + Convert.ToDecimal(bid.Element("C23").Value));  //합계
                        }
                        if (thirdConstruction != null)  //현재 보는 object가 마지막 상위 공종에 포함되어 있다면 단가별 합과 합계를 더해나감
                        {
                            thirdConstruction.Element("C20").Value = string.Concat(Convert.ToDecimal(thirdConstruction.Element("C20").Value) + Convert.ToDecimal(bid.Element("C20").Value));    //재료비
                            thirdConstruction.Element("C21").Value = string.Concat(Convert.ToDecimal(thirdConstruction.Element("C21").Value) + Convert.ToDecimal(bid.Element("C21").Value));    //노무비
                            thirdConstruction.Element("C22").Value = string.Concat(Convert.ToDecimal(thirdConstruction.Element("C22").Value) + Convert.ToDecimal(bid.Element("C22").Value));    //경비
                            thirdConstruction.Element("C23").Value = string.Concat(Convert.ToDecimal(thirdConstruction.Element("C23").Value) + Convert.ToDecimal(bid.Element("C23").Value));    //합계 
                        }
                    }
                }
            }
        }

        static void SubstitutePrice()
        {  //BID 파일 내 원가계산서 관련 금액 세팅
            foreach (var bid in eleBID)
            {
                if (bid.Name == "T5")   //bid.Name이 T5인지를 확인함으로 간단하게 원가 계산서부분의 element 인지를 판별. Tag는 T3가 아닌 T5 기준을 따른다. (23.01.31 수정)
                {
                    if (Data.Bidding.ContainsKey(string.Concat(bid.Element("C4").Value)))
                    {
                        bid.Element("C8").Value = Data.Bidding[string.Concat(bid.Element("C4").Value)].ToString();
                    }
                    else if (Data.Rate1.ContainsKey(string.Concat(bid.Element("C4").Value)))
                    {
                        bid.Element("C8").Value = Data.Bidding[string.Concat(bid.Element("C4").Value)].ToString();
                    }
                }
            }

            if(File.Exists(Data.work_path + "\\Result_Xml.xml"))  //기존 Result_Xml 파일은 삭제한다. (23.02.02)
            {
                File.Delete(Data.work_path + "\\Result_Xml.xml");
            }

            //작업후 xml 파일 저장
            StringBuilder sb = new StringBuilder();
            XmlWriterSettings xws = new XmlWriterSettings
            {
                OmitXmlDeclaration = true,
                Indent = true
            };
            using (XmlWriter xw = XmlWriter.Create(sb, xws))
            {
                docBID.WriteTo(xw);
            }
            File.WriteAllText(Path.Combine(Data.work_path, "Result_Xml.xml"), sb.ToString());
        }

        public static void CreateZipFile(IEnumerable<string> files)
        {
            if (File.Exists(Data.work_path + "\\입찰내역.zip"))  //기존 입찰내역.zip 파일은 삭제한다. (23.02.02)
            {
                File.Delete(Data.work_path + "\\입찰내역.zip");
            }

            var Zip = ZipFile.Open(Path.Combine(Data.work_path, "입찰내역.zip"), ZipArchiveMode.Create);
            foreach (var file in files)
            {
                Zip.CreateEntryFromFile(file, Path.GetFileName(file), CompressionLevel.Optimal);
            }
            Zip.Dispose();
        }
        static void CreateFile()
        {
            //최종 입찰내역 파일 세부공사별로 생성 
            CreateResultFile.Create();
            //생성된 입찰내역 파일 압축 
            string[] files = Directory.GetFiles(Data.folder, "*.xls");  //폴더 경로 수정 (23.02.02)
            CreateZipFile(files);
        }
        static void Reset()
        {
            Data.ExecuteReset = "1";    //Reset 함수 사용 여부

            var DM = Data.Investigation["직접재료비"];
            var DL = Data.Investigation["직접노무비"];
            var OE = Data.Investigation["산출경비"];
            var FM = Data.InvestigateFixedPriceDirectMaterial;
            var FL = Data.InvestigateFixedPriceDirectLabor;
            var FOE = Data.InvestigateFixedPriceOutputExpense;
            var SM = Data.InvestigateStandardMaterial;
            var SL = Data.InvestigateStandardLabor;
            var SOE = Data.InvestigateStandardExpense;
            //조사 내역서 정보 백업

            Data.RealDirectMaterial = DM;
            Data.RealDirectLabor = DL;
            Data.RealOutputExpense = OE;
            Data.FixedPriceDirectMaterial = FM;
            Data.FixedPriceDirectLabor = FL;
            Data.FixedPriceOutputExpense = FOE;
            Data.StandardMaterial = SM;
            Data.StandardLabor = SL;
            Data.StandardExpense = SOE;
            //사정율 재적용을 위한 초기화

            foreach (var bid in eleBID) //Dictionary 초기화
            {
                //일반 항목인 경우
                if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")
                {
                    var constNum = string.Concat(bid.Element("C1").Value);      //세부공사 번호
                    var numVal = string.Concat(bid.Element("C2").Value);        //세부공종 번호
                    var detailVal = string.Concat(bid.Element("C3").Value);     //세부 공종 번호

                    //현재 탐색 공종
                    var curObject = Data.Dic[constNum].Find(x => x.WorkNum == numVal && x.DetailWorkNum == detailVal);
                    curObject.MaterialUnit = Convert.ToDecimal(string.Concat(bid.Element("C16").Value));
                    curObject.LaborUnit = Convert.ToDecimal(string.Concat(bid.Element("C17").Value));
                    curObject.ExpenseUnit = Convert.ToDecimal(string.Concat(bid.Element("C18").Value));
                }
            }
            Data.ExecuteReset = "0";    //Reset 함수 사용이 끝나면 다시 0으로 초기화
        }
        public static void Calculation()
        {
            docBID = XDocument.Load(Path.Combine(Data.folder, "Setting_Xml.xml"));  //폴더 경로 수정 (23.02.02)
            eleBID = docBID.Root.Elements();
            //decimal profitAdjustmentLine = (Data.Bidding["이윤"] / Data.ResultPrice * 100) / Convert.ToDecimal(Data.profitPercent);   //이윤 기준율 대비 현재 이윤의 반영비율
            //decimal managementAdjustmentLine = (Data.Bidding["일반관리비"] / Data.ResultPrice * 100) / Convert.ToDecimal(Data.managementPercent);    //일반 관리비 기준율 대비 현재 일반 관리비의 반영비율
            //가격 재세팅 후 리셋 함수 실행 횟수 증가
            Reset();

            CalculateResultPrice(); //입찰 금액 계산
            CalculateEvaluationPrice(); //자제인력 평가 항목 금액 계산

            Recalculation();    //사정율에 따른 재계산

            SetPriceOfSuperConstruction();  //공종 합계 bid에 저장 (23.02.07)

            FillCostAccount.CalculateBiddingCosts();    //원가계산서 사정율적용(입찰) 금액 계산 및 저장
            FillCostAccount.Adjustment();   //차이 값 보정

            SetBusinessInfo();      //사업자등록번호 <T1></C17></T1>에 추가
            SubstitutePrice();      //원가계산서 사정율 적용하여 계산한 금액들 BID 파일에도 반영
            CreateFile();           //입찰내역 파일 생성
        }
    }
}