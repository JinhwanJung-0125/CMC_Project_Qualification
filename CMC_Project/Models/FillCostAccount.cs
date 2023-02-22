using System;
using System.Collections.Generic;
using System.IO;

/*
 23.02.02 업데이트2
 --------------------
 기존 폴더가 존재해도 제대로 작동되도록 수정
 --------------------
*/
/*
 23.02.06 업데이트
 --------------------
 Data.Investigation, Data.Bidding에 작업설 추가
 --------------------
*/

namespace SetUnitPriceByExcel
{
    class FillCostAccount
    {
        //원가계산서 항목별 조사금액 채움(관리자 보정 후)
        public static void FillInvestigationCosts()
        {

            string costStatementPath = null;
            //원가 계산서 양식 불러오기
            var workbook = ExcelHandling.GetWorkbook("세부결과_원가계산서.xlsx", ".xlsx");
            var sheet = workbook.GetSheetAt(0);
            //적용비율1 작성
            ExcelHandling.GetCell(sheet, 7, 6).SetCellValue(Data.Rate1["간접노무비"] + " %");
            ExcelHandling.GetCell(sheet, 10, 6).SetCellValue(Data.Rate1["산재보험료"] + " %");
            ExcelHandling.GetCell(sheet, 11, 6).SetCellValue(Data.Rate1["고용보험료"] + " %");
            ExcelHandling.GetCell(sheet, 19, 6).SetCellValue(Data.Rate1["환경보전비"] + " %");
            ExcelHandling.GetCell(sheet, 20, 6).SetCellValue(Data.Rate1["공사이행보증서발급수수료"] + " %");
            ExcelHandling.GetCell(sheet, 21, 6).SetCellValue(Data.Rate1["건설하도급보증수수료"] + " %");
            ExcelHandling.GetCell(sheet, 22, 6).SetCellValue(Data.Rate1["건설기계대여대금 지급보증서발급금액"] + " %");
            ExcelHandling.GetCell(sheet, 23, 6).SetCellValue(Data.Rate1["기타경비"] + " %");
            ExcelHandling.GetCell(sheet, 24, 6).SetCellValue(Data.Rate1["일반관리비"] + " %");
            ExcelHandling.GetCell(sheet, 29, 6).SetCellValue(Data.Rate1["공사손해보험료"] + " %");
            ExcelHandling.GetCell(sheet, 31, 6).SetCellValue(Data.Rate1["부가가치세"] + " %");

            //적용비율 2 작성
            ExcelHandling.GetCell(sheet, 7, 7).SetCellValue(Data.Rate2["간접노무비"] + " %");
            ExcelHandling.GetCell(sheet, 10, 7).SetCellValue(Data.Rate2["산재보험료"] + " %");
            ExcelHandling.GetCell(sheet, 11, 7).SetCellValue(Data.Rate2["고용보험료"] + " %");
            ExcelHandling.GetCell(sheet, 23, 7).SetCellValue(Data.Rate2["기타경비"] + " %");

            //금액 세팅
            ExcelHandling.GetCell(sheet, 2, 8).SetCellValue(Data.Investigation["순공사원가"]);      //1. 순공사원가
            ExcelHandling.GetCell(sheet, 3, 8).SetCellValue(Data.Investigation["직접재료비"]);      //가. 재료비
            ExcelHandling.GetCell(sheet, 4, 8).SetCellValue(Data.Investigation["직접재료비"]);      //가-1. 직접재료비
            ExcelHandling.GetCell(sheet, 5, 8).SetCellValue(Data.Investigation["노무비"]);         //나. 노무비
            ExcelHandling.GetCell(sheet, 6, 8).SetCellValue(Data.Investigation["직접노무비"]);      //나-1. 직접노무비
            ExcelHandling.GetCell(sheet, 7, 8).SetCellValue(Data.Investigation["간접노무비"]);      //나-2. 간접노무비
            ExcelHandling.GetCell(sheet, 8, 8).SetCellValue(Data.Investigation["경비"]);           //다. 경비
            ExcelHandling.GetCell(sheet, 9, 8).SetCellValue(Data.Investigation["산출경비"]);        //다-1. 산출경비
            ExcelHandling.GetCell(sheet, 10, 8).SetCellValue(Data.Investigation["산재보험료"]);     //다-2. 산재보험료
            ExcelHandling.GetCell(sheet, 11, 8).SetCellValue(Data.Investigation["고용보험료"]);     //다-3. 고용보험료
            ExcelHandling.GetCell(sheet, 12, 8).SetCellValue(Data.Fixed["국민건강보험료"]);      //다-4. 국민건강보험료
            ExcelHandling.GetCell(sheet, 13, 8).SetCellValue(Data.Fixed["노인장기요양보험"]);    //다-5. 노인장기요양보험
            ExcelHandling.GetCell(sheet, 14, 8).SetCellValue(Data.Fixed["국민연금보험료"]);      //다-6. 국민연금보험료
            ExcelHandling.GetCell(sheet, 15, 8).SetCellValue(Data.Fixed["퇴직공제부금"]);       //다-7. 퇴직공제부금
            ExcelHandling.GetCell(sheet, 16, 8).SetCellValue(Data.Fixed["산업안전보건관리비"]);   //다-8. 산업안전보건관리비
            ExcelHandling.GetCell(sheet, 17, 8).SetCellValue(Data.Fixed["안전관리비"]);         //다-9. 안전관리비
            ExcelHandling.GetCell(sheet, 18, 8).SetCellValue(Data.Fixed["품질관리비"]);         //다-10. 품질관리비
            ExcelHandling.GetCell(sheet, 19, 8).SetCellValue(Data.Investigation["환경보전비"]); //다-11. 환경보전비
            ExcelHandling.GetCell(sheet, 20, 8).SetCellValue(Data.Investigation["공사이행보증서발급수수료"]);   //다-12. 공사이행보증수수료
            ExcelHandling.GetCell(sheet, 21, 8).SetCellValue(Data.Investigation["건설하도급보증수수료"]);      //다-13. 하도급대금지급 보증수수료
            ExcelHandling.GetCell(sheet, 22, 8).SetCellValue(Data.Investigation["건설기계대여대금 지급보증서발급금액"]);    //다-14. 건설기계대여대금 지급보증서 발급금액
            ExcelHandling.GetCell(sheet, 23, 8).SetCellValue(Data.Investigation["기타경비"]);               //다-15. 기타경비
            ExcelHandling.GetCell(sheet, 24, 8).SetCellValue(Data.Investigation["일반관리비"]); //2. 일반관리비
            ExcelHandling.GetCell(sheet, 25, 8).SetCellValue(Data.Investigation["이윤"]);  //3. 이윤
            ExcelHandling.GetCell(sheet, 26, 8).SetCellValue(Data.Investigation["PS"]);   //3.1 PS
            ExcelHandling.GetCell(sheet, 27, 8).SetCellValue(Data.Investigation["제요율적용제외공종"]); //3.2 제요율적용제외공종
            ExcelHandling.GetCell(sheet, 28, 8).SetCellValue(Data.Investigation["총원가"]);  //4. 총원가
            ExcelHandling.GetCell(sheet, 29, 8).SetCellValue(Data.Investigation["공사손해보험료"]);  //5. 공사손해보험료
            ExcelHandling.GetCell(sheet, 30, 8).SetCellValue(Data.Investigation["소계"]);  //6. 소계
            ExcelHandling.GetCell(sheet, 31, 8).SetCellValue(Data.Investigation["부가가치세"]);  //7. 부가가치세
            ExcelHandling.GetCell(sheet, 32, 8).SetCellValue(0);  //8. 매입세
            ExcelHandling.GetCell(sheet, 33, 8).SetCellValue(Data.Investigation["도급비계"]);  //9. 도급비계

            //원가계산서 조사금액 세팅 시점에 CalculatePrice.cs에서 재계산 시, 초기화를 위한 조사금액 저장
            var FM = Data.FixedPriceDirectMaterial;
            var FL = Data.FixedPriceDirectLabor;
            var FOE = Data.FixedPriceOutputExpense;
            var SM = Data.StandardMaterial;
            var SL = Data.StandardLabor;
            var SOE = Data.StandardExpense;

            Data.InvestigateFixedPriceDirectMaterial = FM;
            Data.InvestigateFixedPriceDirectLabor = FL;
            Data.InvestigateFixedPriceOutputExpense = FOE;
            Data.InvestigateStandardMaterial = SM;
            Data.InvestigateStandardLabor = SL;
            Data.InvestigateStandardExpense = SOE;

            if(File.Exists(Data.work_path + "원가계산서.xlsx"))  //먼저 기존 원가계산서 파일이 있다면 삭제한다. (23.02.02)
            {
                File.Delete(Data.work_path + "원가계산서.xlsx");
            }

            costStatementPath = Path.Combine(Data.work_path, "원가계산서.xlsx");
            ExcelHandling.WriteExcel(workbook, costStatementPath);
        }

        //원가계산서 항목별 입찰금액 채움
        public static void FillBiddingCosts()
        {
            //조사금액을 채운 원가계산서_세부결과.xlsx의 경로
            string costStatementPath = Path.Combine(Data.work_path, "원가계산서.xlsx");
            //원가계산서_세부결과 파일 불러오기
            var workbook = ExcelHandling.GetWorkbook(costStatementPath, ".xlsx");

            var sheet = workbook.GetSheetAt(0);

            //적용비율 1, 2 적용금액 원가계산서 반영
            ExcelHandling.GetCell(sheet, 7, 9).SetCellValue(Data.Bidding["간접노무비1"]);
            ExcelHandling.GetCell(sheet, 10, 9).SetCellValue(Data.Bidding["산재보험료1"]);
            ExcelHandling.GetCell(sheet, 11, 9).SetCellValue(Data.Bidding["고용보험료1"]);
            ExcelHandling.GetCell(sheet, 23, 9).SetCellValue(Data.Bidding["기타경비1"]);
            ExcelHandling.GetCell(sheet, 7, 10).SetCellValue(Data.Bidding["간접노무비2"]);
            ExcelHandling.GetCell(sheet, 10, 10).SetCellValue(Data.Bidding["산재보험료2"]);
            ExcelHandling.GetCell(sheet, 11, 10).SetCellValue(Data.Bidding["고용보험료2"]);
            ExcelHandling.GetCell(sheet, 23, 10).SetCellValue(Data.Bidding["기타경비2"]);

            //적용비율 1, 2 적용 금액 중, 큰 금액 세팅
            ExcelHandling.GetCell(sheet, 7, 11).SetCellValue(Data.Bidding["간접노무비max"]);
            ExcelHandling.GetCell(sheet, 10, 11).SetCellValue(Data.Bidding["산재보험료max"]);
            ExcelHandling.GetCell(sheet, 11, 11).SetCellValue(Data.Bidding["고용보험료max"]);
            ExcelHandling.GetCell(sheet, 23, 11).SetCellValue(Data.Bidding["기타경비max"]);

            //금액 세팅
            ExcelHandling.GetCell(sheet, 2, 19).SetCellValue(Data.Bidding["순공사원가"]);      //1. 순공사원가
            ExcelHandling.GetCell(sheet, 3, 19).SetCellValue(Data.Bidding["직접재료비"]);      //가. 재료비
            ExcelHandling.GetCell(sheet, 4, 19).SetCellValue(Data.Bidding["직접재료비"]);      //가-1. 직접재료비
            ExcelHandling.GetCell(sheet, 5, 19).SetCellValue(Data.Bidding["노무비"]);         //나. 노무비
            ExcelHandling.GetCell(sheet, 6, 19).SetCellValue(Data.Bidding["직접노무비"]);      //나-1. 직접노무비
            ExcelHandling.GetCell(sheet, 7, 19).SetCellValue(Data.Bidding["간접노무비"]);      //나-2. 간접노무비
            ExcelHandling.GetCell(sheet, 8, 19).SetCellValue(Data.Bidding["경비"]);           //다. 경비
            ExcelHandling.GetCell(sheet, 9, 19).SetCellValue(Data.Bidding["산출경비"]);        //다-1. 산출경비
            ExcelHandling.GetCell(sheet, 10, 19).SetCellValue(Data.Bidding["산재보험료"]);     //다-2. 산재보험료
            ExcelHandling.GetCell(sheet, 11, 19).SetCellValue(Data.Bidding["고용보험료"]);     //다-3. 고용보험료
            ExcelHandling.GetCell(sheet, 12, 19).SetCellValue(Data.Fixed["국민건강보험료"]);      //다-4. 국민건강보험료
            ExcelHandling.GetCell(sheet, 13, 19).SetCellValue(Data.Fixed["노인장기요양보험"]);    //다-5. 노인장기요양보험
            ExcelHandling.GetCell(sheet, 14, 19).SetCellValue(Data.Fixed["국민연금보험료"]);      //다-6. 국민연금보험료
            ExcelHandling.GetCell(sheet, 15, 19).SetCellValue(Data.Fixed["퇴직공제부금"]);       //다-7. 퇴직공제부금
            ExcelHandling.GetCell(sheet, 16, 19).SetCellValue(Data.Fixed["산업안전보건관리비"]);   //다-8. 산업안전보건관리비
            ExcelHandling.GetCell(sheet, 17, 19).SetCellValue(Data.Fixed["안전관리비"]);         //다-9. 안전관리비
            ExcelHandling.GetCell(sheet, 18, 19).SetCellValue(Data.Fixed["품질관리비"]);         //다-10. 품질관리비
            ExcelHandling.GetCell(sheet, 19, 19).SetCellValue(Data.Bidding["환경보전비"]); //다-11. 환경보전비
            ExcelHandling.GetCell(sheet, 20, 19).SetCellValue(Data.Bidding["공사이행보증서발급수수료"]);   //다-12. 공사이행보증수수료
            ExcelHandling.GetCell(sheet, 21, 19).SetCellValue(Data.Bidding["건설하도급보증수수료"]);      //다-13. 하도급대금지급 보증수수료
            ExcelHandling.GetCell(sheet, 22, 19).SetCellValue(Data.Bidding["건설기계대여대금 지급보증서발급금액"]);    //다-14. 건설기계대여대금 지급보증서 발급금액
            ExcelHandling.GetCell(sheet, 23, 19).SetCellValue(Data.Bidding["기타경비"]);               //다-15. 기타경비
            ExcelHandling.GetCell(sheet, 24, 19).SetCellValue(Data.Bidding["일반관리비"]); //2. 일반관리비
            ExcelHandling.GetCell(sheet, 25, 19).SetCellValue(Data.Bidding["이윤"]);  //3. 이윤
            ExcelHandling.GetCell(sheet, 26, 19).SetCellValue(Data.Bidding["PS"]);   //3.1 PS
            ExcelHandling.GetCell(sheet, 27, 19).SetCellValue(Data.Bidding["제요율적용제외공종"]); //3.2 제요율적용제외공종
            ExcelHandling.GetCell(sheet, 28, 19).SetCellValue(Data.Bidding["총원가"]);  //4. 총원가
            ExcelHandling.GetCell(sheet, 29, 19).SetCellValue(Data.Bidding["공사손해보험료"]);  //5. 공사손해보험료
            ExcelHandling.GetCell(sheet, 30, 19).SetCellValue(Data.Bidding["소계"]);  //6. 소계
            ExcelHandling.GetCell(sheet, 31, 19).SetCellValue(Data.Bidding["부가가치세"]);  //7. 부가가치세
            ExcelHandling.GetCell(sheet, 32, 19).SetCellValue(0);  //8. 매입세
            ExcelHandling.GetCell(sheet, 33, 19).SetCellValue(Data.Bidding["도급비계"]);  //9. 도급비계

            //비율 세팅
            ExcelHandling.GetCell(sheet, 4, 20).SetCellValue((double)GetRate("직접재료비") + "%");      //가-1. 직접재료비
            ExcelHandling.GetCell(sheet, 6, 20).SetCellValue((double)GetRate("직접노무비") + " %");      //나-1. 직접노무비
            ExcelHandling.GetCell(sheet, 7, 20).SetCellValue((double)GetRate("간접노무비") + " %");      //나-2. 간접노무비
            ExcelHandling.GetCell(sheet, 9, 20).SetCellValue((double)GetRate("산출경비") + " %");        //다-1. 산출경비
            ExcelHandling.GetCell(sheet, 10, 20).SetCellValue((double)GetRate("산재보험료") + " %");     //다-2. 산재보험료
            ExcelHandling.GetCell(sheet, 11, 20).SetCellValue((double)GetRate("고용보험료") + " %");     //다-3. 고용보험료
            ExcelHandling.GetCell(sheet, 12, 20).SetCellValue((double)GetRate("국민건강보험료") + " %");      //다-4. 국민건강보험료
            ExcelHandling.GetCell(sheet, 13, 20).SetCellValue((double)GetRate("노인장기요양보험") + " %");    //다-5. 노인장기요양보험
            ExcelHandling.GetCell(sheet, 14, 20).SetCellValue((double)GetRate("국민연금보험료") + " %");      //다-6. 국민연금보험료
            ExcelHandling.GetCell(sheet, 15, 20).SetCellValue((double)GetRate("퇴직공제부금") + " %");       //다-7. 퇴직공제부금
            ExcelHandling.GetCell(sheet, 16, 20).SetCellValue((double)GetRate("산업안전보건관리비") + " %");   //다-8. 산업안전보건관리비
            ExcelHandling.GetCell(sheet, 17, 20).SetCellValue((double)GetRate("안전관리비") + " %");         //다-9. 안전관리비
            ExcelHandling.GetCell(sheet, 18, 20).SetCellValue((double)GetRate("품질관리비") + " %");         //다-10. 품질관리비
            ExcelHandling.GetCell(sheet, 19, 20).SetCellValue((double)GetRate("환경보전비") + " %");         //다-11. 환경보전비
            ExcelHandling.GetCell(sheet, 20, 20).SetCellValue((double)GetRate("공사이행보증서발급수수료") + " %");   //다-12. 공사이행보증수수료
            ExcelHandling.GetCell(sheet, 21, 20).SetCellValue((double)GetRate("건설하도급보증수수료") + " %");      //다-13. 하도급대금지급 보증수수료
            ExcelHandling.GetCell(sheet, 22, 20).SetCellValue((double)GetRate("건설기계대여대금 지급보증서발급금액") + " %");    //다-14. 건설기계대여대금 지급보증서 발급금액
            ExcelHandling.GetCell(sheet, 23, 20).SetCellValue((double)GetRate("기타경비") + " %");               //다-15. 기타경비
            ExcelHandling.GetCell(sheet, 24, 20).SetCellValue((double)GetRate("일반관리비") + " %"); //2. 일반관리비
            ExcelHandling.GetCell(sheet, 25, 20).SetCellValue("0%");  //3. 이윤
            ExcelHandling.GetCell(sheet, 26, 20).SetCellValue((double)GetRate("PS") + " %");   //3.1 PS
            ExcelHandling.GetCell(sheet, 27, 20).SetCellValue((double)GetRate("제요율적용제외공종") + " %"); //3.2 제요율적용제외공종
            ExcelHandling.GetCell(sheet, 29, 20).SetCellValue((double)GetRate("공사손해보험료") + " %"); //5. 공사손해보험료
            ExcelHandling.GetCell(sheet, 33, 20).SetCellValue((double)GetRate("도급비계") + " %");  //9. 도급비계

            costStatementPath = Path.Combine(Data.work_path, "원가계산서_세부결과.xlsx");
            ExcelHandling.WriteExcel(workbook, costStatementPath);
        }

        //원가계산서 항목별 조사금액 구하여 Dictionary Investigation에 저장
        //보정의 경우, 매개변수로 보정할 항목의 이름(item)과 보정할 금액(price)를 받아 값을 적용
        public static void CalculateInvestigationCosts(Dictionary<string, long> correction)
        {
            //직공비
            Data.Investigation["직공비"] = ToLong(Data.RealDirectMaterial + Data.RealDirectLabor + Data.RealOutputExpense);
            //가-1. 직접재료비
            Data.Investigation["직접재료비"] = ToLong(Data.RealDirectMaterial);
            //나-1. 직접노무비
            Data.Investigation["직접노무비"] = ToLong(Data.RealDirectLabor);
            //나-2.간접노무비
            Data.Investigation["간접노무비"] = ToLong(Data.RealDirectLabor * (Data.Rate1["간접노무비"] * 0.01m));
            //나. 노무비
            Data.Investigation["노무비"] = ToLong(Data.RealDirectLabor) + Data.Investigation["간접노무비"];
            //다-1. 산출경비
            Data.Investigation["산출경비"] = ToLong(Data.RealOutputExpense);
            //다-2. 산재보험료
            Data.Investigation["산재보험료"] = ToLong(Data.Investigation["노무비"] * (Data.Rate1["산재보험료"] * 0.01m));
            //다-3. 고용보험료
            Data.Investigation["고용보험료"] = ToLong(Data.Investigation["노무비"] * (Data.Rate1["고용보험료"] * 0.01m));
            //다-11. 환경보전비
            Data.Investigation["환경보전비"] = ToLong(Data.Investigation["직공비"] * (Data.Rate1["환경보전비"] * 0.01m));
            //다-12 공사이행보증수수료
            Data.Investigation["공사이행보증서발급수수료"] = 0;
            if (Data.Rate1["공사이행보증서발급수수료"] != 0)
                Data.Investigation["공사이행보증서발급수수료"] = GetConstructionGuaranteeFee(Data.Investigation["직공비"]);
            if (correction.ContainsKey("공사이행보증서발급수수료"))
                Data.Investigation["공사이행보증서발급수수료"] = correction["공사이행보증서발급수수료"];
            //다-13. 하도급대금지금보증수수료
            Data.Investigation["건설하도급보증수수료"] = ToLong(Data.Investigation["직공비"] * (Data.Rate1["건설하도급보증수수료"] * 0.01m));
            //다-14. 건설기계대여대금 지급보증서 발급금액
            Data.Investigation["건설기계대여대금 지급보증서발급금액"] = ToLong(Data.Investigation["직공비"] * (Data.Rate1["건설기계대여대금 지급보증서발급금액"] * 0.01m));
            //다-15. 기타경비
            Data.Investigation["기타경비"] = ToLong((Data.Investigation["직접재료비"] + Data.Investigation["노무비"]) * (Data.Rate1["기타경비"] * 0.01m));
            //다. 경비
            Data.Investigation["경비"] = Data.Investigation["산출경비"] + Data.Investigation["산재보험료"] + Data.Investigation["고용보험료"]
                                            + Data.Fixed["국민건강보험료"] + Data.Fixed["노인장기요양보험"] + Data.Fixed["국민연금보험료"]
                                            + Data.Fixed["퇴직공제부금"] + Data.Fixed["산업안전보건관리비"] + Data.Fixed["안전관리비"]
                                            + Data.Fixed["품질관리비"] + Data.Investigation["환경보전비"] + Data.Investigation["공사이행보증서발급수수료"]
                                            + Data.Investigation["건설하도급보증수수료"] + Data.Investigation["건설기계대여대금 지급보증서발급금액"] + Data.Investigation["기타경비"];
            //1. 순공사원가
            Data.Investigation["순공사원가"] = Data.Investigation["직접재료비"] + Data.Investigation["노무비"] + Data.Investigation["경비"];
            //2. 일반관리비
            Data.Investigation["일반관리비"] = ToLong(Data.Investigation["순공사원가"] * (Data.Rate1["일반관리비"] * 0.01m));
            //3. 이윤
            Data.Investigation["이윤"] = ToLong((Data.Investigation["노무비"] + Data.Investigation["경비"] + Data.Investigation["일반관리비"]) * 0.12m);
            if (correction.ContainsKey("이윤"))
                Data.Investigation["이윤"] = correction["이윤"];
            //3.1 미확정설계공종(PS)
            Data.Investigation["PS"] = ToLong(Data.PsMaterial + Data.PsLabor + Data.PsExpense);
            //3.2 제요율적용제외공종
            Data.Investigation["제요율적용제외공종"] = ToLong(Data.ExcludingMaterial + Data.ExcludingLabor + Data.ExcludingExpense);
            var exSum = Data.ExcludingMaterial + Data.ExcludingLabor + Data.ExcludingExpense;
            var exRate2 = Math.Round(exSum / Data.Investigation["직공비"], 5);
            Data.Rate2["제요율적용제외공종"] = exRate2;
            //4. 총원가
            Data.Investigation["총원가"] = Data.Investigation["순공사원가"] + Data.Investigation["일반관리비"] + Data.Investigation["이윤"] + Data.Investigation["PS"] + Data.Investigation["제요율적용제외공종"];
            //5. 공사손해보험료
            Data.Investigation["공사손해보험료"] = ToLong(Data.Investigation["직공비"] * (Data.Rate1["공사손해보험료"] * 0.01m));
            if (correction.ContainsKey("공사손해보험료"))
                Data.Investigation["공사손해보험료"] = correction["공사손해보험료"];
            //작업설 추가 (23.02.06)
            Data.Investigation["작업설 등"] = ToLong(Data.ByProduct);
            //6. 소계
            Data.Investigation["소계"] = Data.Investigation["총원가"] + Data.Investigation["공사손해보험료"] + Data.Investigation["작업설 등"]; //전체 가격 계산에 작업설 추가 (23.02.06)
            //7. 부가가치세
            Data.Investigation["부가가치세"] = ToLong(Data.Investigation["소계"] * (Data.Rate1["부가가치세"] * 0.01m));
            if (correction.ContainsKey("부가가치세"))
                Data.Investigation["부가가치세"] = correction["부가가치세"];
            //8. 매입세(입찰 공사 파일 중, 매입세 있는 공사 없음. 추후 추가할 수 있음)
            //9. 도급비계
            Data.Investigation["도급비계"] = Data.Investigation["소계"] + Data.Investigation["부가가치세"];
        }

        //원가계산서 항목별 입찰금액 구하여 Bidding에 저장
        public static void CalculateBiddingCosts()
        {
            //직공비
            Data.Bidding["직공비"] = ToLong(Data.RealDirectMaterial + Data.RealDirectLabor + Data.RealOutputExpense);
            //적용비율 2를 적용한 금액 계산
            var undirectlabor2 = Data.Bidding["직공비"] * (Data.Rate2["간접노무비"] * 0.01m);
            var industrial2 = Data.Bidding["직공비"] * (Data.Rate2["산재보험료"] * 0.01m);
            var employ2 = Data.Bidding["직공비"] * (Data.Rate2["고용보험료"] * 0.01m);
            var etc2 = Data.Bidding["직공비"] * (Data.Rate2["기타경비"] * 0.01m);

            //적용비율 2를 적용한 금액 저장
            Data.Bidding["간접노무비2"] = ToLong(undirectlabor2);
            Data.Bidding["산재보험료2"] = ToLong(industrial2);
            Data.Bidding["고용보험료2"] = ToLong(employ2);
            Data.Bidding["기타경비2"] = ToLong(etc2);

            //가-1. 직접재료비
            Data.Bidding["직접재료비"] = ToLong(Data.RealDirectMaterial);
            //나-1. 직접노무비
            Data.Bidding["직접노무비"] = ToLong(Data.RealDirectLabor);
            //나-2.간접노무비
            Data.Bidding["간접노무비"] = ToLong(Data.Bidding["직접노무비"] * 0.113m);
            //나. 노무비
            Data.Bidding["노무비"] = Data.Bidding["직접노무비"] + Data.Bidding["간접노무비"];
            //다-1. 산출경비
            Data.Bidding["산출경비"] = ToLong(Data.RealOutputExpense);
            //다-2. 산재보험료
            Data.Bidding["산재보험료1"] = ToLong(Data.Bidding["노무비"] * (Data.Rate1["산재보험료"] * 0.01m));
            if (Data.Bidding["산재보험료1"] < Data.Bidding["산재보험료2"])
            {
                Data.Bidding["산재보험료"] = Data.Bidding["산재보험료2"];
                Data.Bidding["산재보험료max"] = Data.Bidding["산재보험료2"];
            }
            else
            {
                Data.Bidding["산재보험료"] = Data.Bidding["산재보험료1"];
                Data.Bidding["산재보험료max"] = Data.Bidding["산재보험료1"];
            }
            //다-3. 고용보험료
            Data.Bidding["고용보험료1"] = ToLong(Data.Bidding["노무비"] * (Data.Rate1["고용보험료"] * 0.01m));
            if (Data.Bidding["고용보험료1"] < Data.Bidding["고용보험료2"])
            {
                Data.Bidding["고용보험료"] = Data.Bidding["고용보험료2"];
                Data.Bidding["고용보험료max"] = Data.Bidding["고용보험료2"];
            }
            else
            {
                Data.Bidding["고용보험료"] = Data.Bidding["고용보험료1"];
                Data.Bidding["고용보험료max"] = Data.Bidding["고용보험료1"];
            }
            //다-11. 환경보전비
            Data.Bidding["환경보전비"] = ToLong(Data.Bidding["직공비"] * (Data.Rate1["환경보전비"] * 0.01m));
            //다-12 공사이행보증수수료
            Data.Bidding["공사이행보증서발급수수료"] = 0;
            if (Data.Rate1["공사이행보증서발급수수료"] != 0)
                Data.Bidding["공사이행보증서발급수수료"] = GetConstructionGuaranteeFee(Data.Bidding["직공비"]);
            //다-13. 하도급대금지금보증수수료
            Data.Bidding["건설하도급보증수수료"] = ToLong(Data.Bidding["직공비"] * (Data.Rate1["건설하도급보증수수료"] * 0.01m));
            //다-14. 건설기계대여대금 지급보증서 발급금액
            Data.Bidding["건설기계대여대금 지급보증서발급금액"] = ToLong(Data.Bidding["직공비"] * (Data.Rate1["건설기계대여대금 지급보증서발급금액"] * 0.01m));
            //다. 경비
            Data.Bidding["경비"] = Data.Bidding["산출경비"] + Data.Bidding["산재보험료"] + Data.Bidding["고용보험료"]
                                + Data.Fixed["국민건강보험료"] + Data.Fixed["노인장기요양보험"] + Data.Fixed["국민연금보험료"]
                                + Data.Fixed["퇴직공제부금"] + Data.Fixed["산업안전보건관리비"] + Data.Fixed["안전관리비"]
                                + Data.Fixed["품질관리비"] + Data.Bidding["환경보전비"] + Data.Bidding["공사이행보증서발급수수료"]
                                + Data.Bidding["건설하도급보증수수료"] + Data.Bidding["건설기계대여대금 지급보증서발급금액"] + Data.Bidding["기타경비"];
            //1. 순공사원가
            Data.Bidding["순공사원가"] = Data.Bidding["직접재료비"] + Data.Bidding["노무비"] + Data.Bidding["경비"];
            
            //3.1 미확정설계공종(PS)
            Data.Bidding["PS"] = ToLong(Data.PsMaterial + Data.PsLabor + Data.PsExpense);
            //3.2 제요율적용제외공종
            Data.Bidding["제요율적용제외공종"] = ToLong(Data.AdjustedExMaterial + Data.AdjustedExLabor + Data.AdjustedExExpense);
            //4. 총원가
            Data.Bidding["총원가"] = Data.Bidding["순공사원가"] + Data.Bidding["일반관리비"] + Data.Bidding["이윤"]
                                + Data.Bidding["PS"] + Data.Bidding["제요율적용제외공종"];
            //5. 공사손해보험료
            Data.Bidding["공사손해보험료"] = ToLong(Data.Bidding["직공비"] * (Data.Rate1["공사손해보험료"] * 0.01m));
            if (Data.CostAccountDeduction.Equals("1"))
            {
                Data.Bidding["공사손해보험료before"] = ToLong(Data.Bidding["직공비"] * (Data.Rate1["공사손해보험료"] * 0.01m));
                Data.Bidding["공사손해보험료"] = ToLong(Math.Ceiling(Data.Bidding["직공비"] * Data.Rate1["공사손해보험료"] * 0.01m * 0.997m));
            }
            //작업설 추가 (23.02.06)
            Data.Bidding["작업설 등"] = ToLong(Data.ByProduct);
            //6. 소계
            Data.Bidding["소계"] = Data.Bidding["총원가"] + Data.Bidding["공사손해보험료"] + Data.Bidding["작업설 등"]; //전체 가격 계산에 작업설 추가 (23.02.06)
            //7. 부가가치세
            Data.Bidding["부가가치세"] = ToLong(Data.Bidding["소계"] * (Data.Rate1["부가가치세"] * 0.01m));
            //8. 매입세(입찰 공사 파일 중, 매입세 있는 공사 없음. 추후 추가할 수 있음)
            //9. 도급비계
            Data.Bidding["도급비계"] = Data.Bidding["소계"] + Data.Bidding["부가가치세"];
            //도급비계 1000원 단위 절상 옵션 적용시
            if (Data.BidPriceRaise.Equals("1"))
            {
                var raise = 1000 - (Convert.ToDecimal(Data.Bidding["도급비계"]) % 1000);  //1000원 단위 절상
                var addPrice = raise / 1.1m;
                Data.Bidding["도급비계"] = ToLong(Data.Bidding["도급비계"] + raise);
                Data.Bidding["일반관리비"] = ToLong(Convert.ToDecimal(Data.Bidding["일반관리비"]) + addPrice);   //절상에 필요한 가격을 일반관리비에 더해 금액을 맞추어줌

                //일반관리비 증가에 따른 타 금액 변경
                Data.Bidding["소계"] = ToLong(Convert.ToDecimal(Data.Bidding["소계"]) + addPrice);
                Data.Bidding["부가가치세"] = ToLong(Data.Bidding["소계"] * (Data.Rate1["부가가치세"] * 0.01m));
                //계산된 도급비계 금액이 천원 단위가 아닐 경우, 부가세 조정
                var difference = Data.Bidding["도급비계"] - (Data.Bidding["소계"] + Data.Bidding["부가가치세"]);
                Data.Bidding["부가가치세"] = ToLong(Data.Bidding["부가가치세"] + difference);
                Console.WriteLine("차이 : " + difference);
            }
        }

        public static void Adjustment() //계산된 도급비계와 입찰금액의 차이 값을 보정
        {
            decimal differency = (Data.Bidding["도급비계"] - Data.ResultPrice) / 1.1m;   //차이 값 계산 
            decimal profitDiffer = Data.Bidding["이윤"] - (0.9m * Convert.ToDecimal(Data.profitPercent / 100) * Data.ResultPrice);    //이윤 기준율 대비 90% 수준의 차이 값
            decimal managementDiffer = Data.Bidding["일반관리비"] - (0.9m * Convert.ToDecimal(Data.managementPercent / 100) * Data.ResultPrice); //일반관리비 기준율 대비 90% 수준의 차이 값

            //이윤, 일반 관리비, 기타 경비 순으로 차이 값을 보정한다.
            if (profitDiffer > differency)  //현재 차이 값이 이윤 기준율 대비 90% 수준 차이 값 내에서 보정이 안된다면
            {
                Data.Bidding["이윤"] = ToLong(Data.Bidding["이윤"] - profitDiffer);   //이윤을 기준율 대비 90% 수준으로 보정한다.
                differency -= profitDiffer; //보정된 값 만큼 차이 값을 뺀다.
            }
            else       //현재 차이 값이 이윤 기준율 대비 90% 수준 차이 값 내에서 보정이 된다면
            {
                Data.Bidding["이윤"] = ToLong(Data.Bidding["이윤"] - differency);   //차이 값을 모두 이윤에 보정한다.
                differency = 0; //차이 값을 모두 보정했기 때문에 값은 0이 된다.
            }

            if (managementDiffer > differency)  //현재 차이 값이 일반 관리비 기준율 대비 90% 수준 차이 값 내에서 보정이 안된다면
            {
                Data.Bidding["일반관리비"] = ToLong(Data.Bidding["일반관리비"] - managementDiffer);   //일반 관리비를 기준율 대비 90% 수준으로 보정한다.
                differency -= managementDiffer; //보정된 값 만큼 차이 값을 뺀다.
            }
            else       //현재 차이 값이 일반 관리비 기준율 대비 90% 수준 차이 값 내에서 보정이 된다면
            {
                Data.Bidding["일반관리비"] = ToLong(Data.Bidding["일반관리비"] - differency);   //차이 값을 모두 이윤에 보정한다.
                differency = 0; //차이 값을 모두 보정했기 때문에 값은 0이 된다.
            }

            if (differency != 0) //아직 차이 값이 남아 있다면
            {
                Data.Bidding["기타경비"] = ToLong(Data.Bidding["기타경비"] - differency);   //남은 차이 값을 모두 기타 경비에 보정한다.
            }

            CalculateBiddingCosts();    //보정된 값으로 다시 금액 계산 및 저장

            differency = (Data.Bidding["도급비계"] - Data.ResultPrice) / 1.1m;   //차이 값 다시 계산 

            if (differency >= -1 && differency < 0) //차이 값이 1원 차이면 부가가치세에 보정
            {
                Data.Bidding["부가가치세"] += 1;
                Data.Bidding["도급비계"] = Data.Bidding["소계"] + Data.Bidding["부가가치세"];
            }
        }

        //decimal 금액 원 단위 절사
        public static long ToLong(decimal price)
        {
            return Convert.ToInt64(Math.Truncate(price));
        }

        //공사이행보증서발급수수료 금액 계산 후 반환
        static long GetConstructionGuaranteeFee(long directSum)
        {
            long guaranteeFee = 0;
            decimal rate = Data.Rate1["공사이행보증서발급수수료"] * 0.01m;
            decimal term = Data.ConstructionTerm / 365.0m;
            if (directSum < 7000000000)
                guaranteeFee = ToLong(directSum * rate * term);
            else if (directSum < 12000000000)
                guaranteeFee = ToLong(((directSum - 5000000000) * rate + 2000000) * term);
            else if (directSum < 25000000000)
                guaranteeFee = ToLong(((directSum - 14000000000) * rate + 4000000) * term);
            else if (directSum < 50000000000)
                guaranteeFee = ToLong(((directSum - 25000000000) * rate + 6000000) * term);
            else
                guaranteeFee = ToLong(((directSum - 50000000000) * rate + 10000000) * term);
            return guaranteeFee;
        }

        //입찰 금액의 조사금액 대 비율 저장
        public static decimal GetRate(string item)
        {
            if (Data.Fixed.ContainsKey(item))
                return 100m;
            if (Data.Investigation[item] == 0 && Data.Bidding[item] == 0)
                return 100m;
            //원가계산제경비 옵션 적용 항목은 적용 전, 후의 비율 출력
            if (item.Equals("간접노무비") || item.Equals("기타경비") || item.Equals("일반관리비") || item.Equals("공사손해보험료"))
            {
                string before = item + "before";
                return Math.Round(Convert.ToDecimal(Data.Bidding[item]) / Data.Bidding[before], 7) * 100;
            }
            decimal rate = Math.Round(Convert.ToDecimal(Data.Bidding[item]) / Data.Investigation[item], 7);
            rate = rate * 100;
            return rate;
        }

        //해당 공사에 특정 원가계산서 항목이 존재하지 않는 경우
        public static void CheckKeyNotFound()
        {
            if (!Data.Rate1.ContainsKey("간접노무비"))
                Data.Rate1["간접노무비"] = 0;
            if (!Data.Rate1.ContainsKey("산재보험료"))
                Data.Rate1["산재보험료"] = 0;
            if (!Data.Rate1.ContainsKey("고용보험료"))
                Data.Rate1["고용보험료"] = 0;
            if (!Data.Rate1.ContainsKey("환경보전비"))
                Data.Rate1["환경보전비"] = 0;
            if (!Data.Rate1.ContainsKey("공사이행보증서발급수수료"))
                Data.Rate1["공사이행보증서발급수수료"] = 0;
            if (!Data.Rate1.ContainsKey("건설하도급보증수수료"))
                Data.Rate1["건설하도급보증수수료"] = 0;
            if (!Data.Rate1.ContainsKey("건설기계대여대금 지급보증서발급금액"))
                Data.Rate1["건설기계대여대금 지급보증서발급금액"] = 0;
            if (!Data.Rate1.ContainsKey("기타경비"))
                Data.Rate1["기타경비"] = 0;
            if (!Data.Rate1.ContainsKey("일반관리비"))
                Data.Rate1["일반관리비"] = 0;
            if (!Data.Rate1.ContainsKey("부가가치세"))
                Data.Rate1["부가가치세"] = 0;
            if (!Data.Rate1.ContainsKey("공사손해보험료"))
                Data.Rate1["공사손해보험료"] = 0;

            if (!Data.Fixed.ContainsKey("국민건강보험료"))
                Data.Fixed["국민건강보험료"] = 0;
            if (!Data.Fixed.ContainsKey("노인장기요양보험"))
                Data.Fixed["노인장기요양보험"] = 0;
            if (!Data.Fixed.ContainsKey("국민연금보험료"))
                Data.Fixed["국민연금보험료"] = 0;
            if (!Data.Fixed.ContainsKey("퇴직공제부금"))
                Data.Fixed["퇴직공제부금"] = 0;
            if (!Data.Fixed.ContainsKey("산업안전보건관리비"))
                Data.Fixed["산업안전보건관리비"] = 0;
            if (!Data.Fixed.ContainsKey("안전관리비"))
                Data.Fixed["안전관리비"] = 0;
            if (!Data.Fixed.ContainsKey("품질관리비"))
                Data.Fixed["품질관리비"] = 0;

            if (!Data.Rate2.ContainsKey("간접노무비"))
                Data.Rate2["간접노무비"] = 0;
            if (!Data.Rate2.ContainsKey("산재보험료"))
                Data.Rate2["산재보험료"] = 0;
            if (!Data.Rate2.ContainsKey("고용보험료"))
                Data.Rate2["고용보험료"] = 0;
            if (!Data.Rate2.ContainsKey("기타경비"))
                Data.Rate2["기타경비"] = 0;
        }
    }
}