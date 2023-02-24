using Microsoft.Win32;
using SetUnitPriceByExcel;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
/*
 23.02.01 업데이트
 --------------------
  사정율에 음의 값 입력 가능하도록 수정
  사정율, 사업자등록번호 값 검토 추가
 --------------------
 23.01.31 업데이트
 --------------------
  UI 수정
  사업자등록번호 입력받기&Data에 저장
  BusinessChangedHandler() 추가
  SetBusinessInfoBtnClick() 추가
 --------------------
*/
namespace CMC_Project.Views
{
    /// <summary>
    /// Interaction logic for AdjustmentPage.xaml
    /// </summary>
    public partial class AdjustmentPage : Page
    {
        private static bool isCalculate = false;
        public static bool isConfirm = false;
        public AdjustmentPage()
        {
            InitializeComponent();
            this.businessNum.TextChanged += BusinessNumChangedHandler;
            this.businessName.TextChanged += BusinessNameChangedHandler;
            this.priceA.TextChanged += PriceAChangedHandler;
            this.estimateRating.TextChanged += EstimateChangedHandler;
            this.expencePercent.TextChanged += ExpencePercentChangedHandler;
            this.managementPercent.TextChanged += ManagementPercentChangedHandler;
            this.profitPercent.TextChanged += ProfitPercentChangedHandler;


            // 사정율 초기화
            if (Data.priceA != null && Data.PersonalRateNum != null)
            {
                priceA.Text = ((double)Data.priceA).ToString();
                estimateRating.Text = ((double)Data.PersonalRateNum).ToString();
            }
            // 라디오 버튼 초기화
            Data.UnitPriceTrimming = "1";
            // 표준시장 단가 체크
            if (Data.StandardMarketDeduction == "1")
                CheckStandardPrice.IsChecked = true;
            else
                CheckStandardPrice.IsChecked = false;
            // 공종 가중치 체크
            if (Data.ZeroWeightDeduction == "1")
                CheckWeightValue.IsChecked = true;
            else
                CheckWeightValue.IsChecked = false;
            // 법정 요율 체크
            if (Data.CostAccountDeduction == "1")
                CheckCAD.IsChecked = true;
            else
                CheckCAD.IsChecked = false;
            // 원단위 체크
            if (Data.BidPriceRaise == "1")
                CheckCeiling.IsChecked = true;
            else
                CheckCeiling.IsChecked = false;


        }

        //sender: 이벤트 발생자, args: 이벤트 인자

        private void BusinessNumChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox BusinessNum = sender as TextBox;
            BusinessNum.MaxLength = 10;
            int selectionStart = BusinessNum.SelectionStart;
            string result = string.Empty;
            //Data.CompanyRegistrationNum = (Double.Parse(BusinessNum.GetLineText(0)));
            foreach (char character in BusinessNum.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character))
                {
                    result += character;

                }
            }
            BusinessNum.Text = result;
            BusinessNum.SelectionStart = selectionStart <= BusinessNum.Text.Length ? selectionStart : BusinessNum.Text.Length;
        }

        private void BusinessNameChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox BusinessName = sender as TextBox;
            int selectionStart = BusinessName.SelectionStart;
            string result = string.Empty;
            Data.CompanyRegistrationName = BusinessName.GetLineText(0);
            foreach (char character in BusinessName.Text.ToCharArray())
            {
                    result += character;
            }
            BusinessName.Text = result;
            BusinessName.SelectionStart = selectionStart <= BusinessName.Text.Length ? selectionStart : BusinessName.Text.Length;
        }


        private void PriceAChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox priceA = sender as TextBox;
            int selectionStart = priceA.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;
            //Data.BalanceRateNum = (Double.Parse(averageRating.GetLineText(0)));

            foreach (char character in priceA.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0) 
                    || (character =='-' && mCount == 0 ))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount +=1;
                    }

                }
            }
            priceA.Text = result;
            priceA.SelectionStart = selectionStart <= priceA.Text.Length ? selectionStart : priceA.Text.Length;
        }

        private void BasePriceChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox basePrice = sender as TextBox;
            int selectionStart = basePrice.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;
            //Data.BalanceRateNum = (Double.Parse(averageRating.GetLineText(0)));

            foreach (char character in basePrice.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0)
                    || (character == '-' && mCount == 0))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount += 1;
                    }

                }
            }
            basePrice.Text = result;
            basePrice.SelectionStart = selectionStart <= basePrice.Text.Length ? selectionStart : basePrice.Text.Length;
        }

        private void ResultPriceChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox resultPrice = sender as TextBox;
            int selectionStart = resultPrice.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;
            //Data.BalanceRateNum = (Double.Parse(averageRating.GetLineText(0)));

            foreach (char character in resultPrice.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0)
                    || (character == '-' && mCount == 0))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount += 1;
                    }

                }
            }
            resultPrice.Text = result;
            resultPrice.SelectionStart = selectionStart <= resultPrice.Text.Length ? selectionStart : resultPrice.Text.Length;
        }

        //sender: 이벤트 발생자, args: 이벤트 인자
        private void EstimateChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox estimateRating = sender as TextBox;
            int selectionStart = estimateRating.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;
            //Data.PersonalRateNum = (Double.Parse(estimateRating.GetLineText(0)));


            foreach (char character in estimateRating.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0)
                    || (character == '-' && mCount == 0))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount += 1;
                    }
                }
            }
            
            estimateRating.Text = result;
            estimateRating.SelectionStart = selectionStart <= estimateRating.Text.Length ? selectionStart : estimateRating.Text.Length;
        }

        private void LaborPercentChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox laborPercent = sender as TextBox;
            int selectionStart = laborPercent.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;

            foreach(char character in laborPercent.Text.ToCharArray())
            {
                if(char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0) || (character == '-' && mCount == 0))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount += 1;
                    }
                }
            }

            laborPercent.Text = result;
            laborPercent.SelectionStart = selectionStart <= laborPercent.Text.Length ? selectionStart : laborPercent.Text.Length;
        }

        private void ExpencePercentChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox expencePercent = sender as TextBox;
            int selectionStart = expencePercent.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;

            foreach (char character in expencePercent.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0) || (character == '-' && mCount == 0))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount += 1;
                    }
                }
            }

            expencePercent.Text = result;
            expencePercent.SelectionStart = selectionStart <= expencePercent.Text.Length ? selectionStart : expencePercent.Text.Length;
        }

        private void ManagementPercentChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox managementPercent = sender as TextBox;
            int selectionStart = managementPercent.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;

            foreach (char character in managementPercent.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0) || (character == '-' && mCount == 0))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount += 1;
                    }
                }
            }

            managementPercent.Text = result;
            managementPercent.SelectionStart = selectionStart <= managementPercent.Text.Length ? selectionStart : managementPercent.Text.Length;
        }

        private void ProfitPercentChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox profitPercent = sender as TextBox;
            int selectionStart = profitPercent.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;

            foreach (char character in profitPercent.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0) || (character == '-' && mCount == 0))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount += 1;
                    }
                }
            }

            profitPercent.Text = result;
            profitPercent.SelectionStart = selectionStart <= profitPercent.Text.Length ? selectionStart : profitPercent.Text.Length;
        }

        private void UpBtnClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Clicked");
        }


        // ------------------------- 옵션 입력 버튼 ------------------------------------------------------------------------------------------------------------------------------------------- //
        //소수 1자리 체크
        private void RadioDecimal_Checked(object sender, RoutedEventArgs e)
        {
            Data.UnitPriceTrimming = "1";
        }
        // 정수 체크
        private void RadioInteger_Checked(object sender, RoutedEventArgs e)
        {
            Data.UnitPriceTrimming = "2";
        }

        // 표준시장 단가 체크
        private void CheckStandardPrice_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckStandardPrice.IsChecked)
            {
                Data.StandardMarketDeduction = "1";
            }
            else
            {
                Data.StandardMarketDeduction = "2";
            }
        }

        // 공종 가중치 체크
        private void CheckWeightValue_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckWeightValue.IsChecked)
            {
                Data.ZeroWeightDeduction = "1";
            }
            else
            {
                Data.ZeroWeightDeduction = "2";
            }
        }

        // 법정요율 체크
        private void CheckCAD_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckCAD.IsChecked)
            {
                Data.CostAccountDeduction = "1";
            }
            else
            {
                Data.CostAccountDeduction = "2";
            }
        }

        // 원단위 체크
        private void CheckCeiling_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckCeiling.IsChecked)
            {
                Data.BidPriceRaise = "1";
            }
            else
            {
                Data.BidPriceRaise = "2";
            }
        }

        //노무비 하한율 체크
        private void CheckLaborCost_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckLaborCost.IsChecked)
            {
                Data.LaborCostLowBound = "1";
            }
            else
            {
                Data.LaborCostLowBound = "2";
            }
        }

        private void SetBusinessInfoBtnClick(object sender, RoutedEventArgs e)
        {

            if (businessNum.Text == string.Empty)
            {
                DisplayDialog("사업자등록번호를 입력해주세요.","Fail");
            }
            else if (businessNum.GetLineLength(0) != 10) // 입력한 사용자등록번호가 10자리가 아닐 때
            {
                DisplayDialog("올바른 사용자등록번호를 입력해주세요.","Fail");
            }
            else if (businessName.Text == string.Empty)
            {
                DisplayDialog("회사명을 입력해주세요.", "Fail");
            }
            else
            {
                Data.CompanyRegistrationName = businessName.Text;
                Data.CompanyRegistrationNum = businessNum.Text;
                DisplayDialog($"입찰업체정보를 저장했습니다.", "Success");
            }
        }

        private void CalBtnClick(object sender, RoutedEventArgs e)
        {
            if (priceA.Text == string.Empty || estimateRating.Text == string.Empty)
            {
                DisplayDialog("A값을 입력해주세요.", "Error");
                return;
            }
            else
            {
                try
                {
                    Data.priceA = (Decimal.Parse(priceA.Text));
                    if (Data.priceA <= 0)
                        throw new Exception();
                }
                catch (Exception)
                {
                    DisplayDialog("A값이 올바르지 않습니다.", "Error");
                    return;
                }
                try
                {
                    Data.BasePrice = (Decimal.Parse(basePrice.Text));
                    if (Data.BasePrice <= 0)
                        throw new Exception();
                }
                catch (Exception)
                {
                    DisplayDialog("기초가격이 올바르지 않습니다.", "Error");
                    return;
                }
                try
                {
                    Data.PersonalRateNum = (Double.Parse(estimateRating.Text));
                    if (Data.PersonalRateNum < -3 || Data.PersonalRateNum > 3)
                        throw new Exception();
                }
                catch (Exception)
                {
                    DisplayDialog("나의 사정율이 올바르지 않습니다.", "Error");
                    return;
                }
                try
                {
                    Data.ResultPrice = (Decimal.Parse(resultPrice.Text));
                    if (Data.ResultPrice < 0)
                        throw new Exception();
                    else if(Data.ResultPrice > 0)
                    {
                        if(Data.ResultPrice > 0.88m * (Data.BasePrice * 1.03m - Data.priceA) + Data.priceA || Data.ResultPrice < 0.88m * (Data.BasePrice * 0.97m - Data.priceA) + Data.priceA)  //직접 입력받은 입찰 금액이 예가 * 88%의 가능 범위(-3% ~ +3%)를 초과하면 에러
                        {
                            throw new Exception();
                        }
                    }
                }
                catch (Exception)
                {
                    DisplayDialog("입찰금액이 올바르지 않습니다.", "Error");
                    return;
                }
               
            }
            if (Data.CompanyRegistrationNum == "")
            {
                DisplayDialog("사업자등록번호를 입력해주세요.", "Error");
                return;
            }
            if(laborPercent.Text == string.Empty || expencePercent.Text == string.Empty || managementPercent.Text == string.Empty || profitPercent.Text == string.Empty) 
            {
                DisplayDialog("기준율을 입력해주세요.", "Error");
                return;
            }
            else
            {
                try
                {
                    Data.laborPercent = (Double.Parse(laborPercent.Text));
                    if(Data.laborPercent > 100 || Data.laborPercent < 0)
                        throw new Exception();
                }
                catch (Exception)
                {
                    DisplayDialog("노무비 기준율이 올바르지 않습니다.", "Error");
                    return;
                }
                try
                {
                    Data.expencePercent = (Double.Parse(expencePercent.Text));
                    if (Data.expencePercent > 100 || Data.expencePercent < 0)
                        throw new Exception();
                }
                catch (Exception)
                {
                    DisplayDialog("기타 경비 기준율이 올바르지 않습니다.", "Error");
                    return;
                }
                try
                {
                    Data.managementPercent = (Double.Parse(managementPercent.Text));
                    if (Data.managementPercent > 100 || Data.managementPercent < 0)
                        throw new Exception();
                }
                catch (Exception)
                {
                    DisplayDialog("일반관리비 기준율이 올바르지 않습니다.", "Error");
                    return;
                }
                try
                {
                    Data.profitPercent = (Double.Parse(profitPercent.Text));
                    if (Data.profitPercent > 100 || Data.profitPercent < 0)
                        throw new Exception();
                }
                catch (Exception)
                {
                    DisplayDialog("이윤 기준율이 올바르지 않습니다.", "Error");
                    return;
                }
            }



            // 단가를 불러온 경우
            if (isConfirm)
            {
                //입찰금액 심사 점수 계산 및 단가 조정
                CalculatePrice.Calculation();

                FixedPercentPrice.Text = Data.FixedPricePercent + " %";
                TargetRate.Text = Data.Bidding["도급비계"] + " 원 " + "(" + FillCostAccount.GetRate("도급비계") + " %)"; // 도급비계
                isCalculate = true;

                //OutputTextBlock.Text = "사정율 적용 완료!";
                DisplayDialog("사정율 적용을 완료하였습니다", "Success");
            }

            // 단가를 불러오지 않은 경우
            else
            {
                DisplayDialog("단가를 먼저 세팅해주세요.", "Error");
            }
        }


        // ------------------------- 세부 결과 확인 버튼 ------------------------------------------------------------------------------------------------------------------------------------------------- //
        private void ShowResult_Click(object sender, RoutedEventArgs e)
        {
            if (isCalculate)
            {
                CMC_Project.Views.ResultPage rw = new();

                rw.Show();
            }
            else
            {
                DisplayDialog("계산 후 확인해주세요", "Fail");
            }
        }



        // 메세지 창
        static public void DisplayDialog(String dialog, String title)
        {
            MessageBox.Show(dialog, title, MessageBoxButton.OK, MessageBoxImage.Information);
        }


        // ------------------------- BID파일 저장 버튼 ---------------------------------------------------------------------------------------------------------------------------------------- //
        private void SaveBidBtnClick(object sender, System.EventArgs e)
        {
            // TargetRate가 계산 되어 있을 경우
            if (isCalculate)
            {
                //단가 세팅 완료한 xml 파일을 다시 BID 파일로 변환
                BidHandling.XmlToBid();

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                saveFileDialog.Filter = "BID Files (*.BID)|*.BID|All files (*.*)|*.*";
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = BidHandling.filename.Substring(0, 16);
                saveFileDialog.OverwritePrompt = true;


                if (saveFileDialog.ShowDialog() == true)
                {
                    string file = saveFileDialog.FileName.ToString(); //경로와 파일명 저장
                    string bidFolder = Data.work_path; //Result Bid 경로
                    string finalBidFile = Path.Combine(bidFolder, BidHandling.filename.Substring(0, 16) + ".BID");

                    File.Move(finalBidFile, file);
                    DisplayDialog("저장되었습니다.", "Save");
                }
                else
                {
                    DisplayDialog("취소되었습니다.", "Error");
                }
            }

            // 계산 안되어 있을 경우
            else
            {
                DisplayDialog("입찰점수를 계산해주세요.", "Error");
            }
        }


        // ------------------------- 원가계산서 저장 버튼 ------------------------------------------------------------------------------------------------------------------------------------- //
        private void SaveCostBtnClick(object sender, System.EventArgs e)
        {
            // TargetRate가 계산 되어 있을 경우
            if (isCalculate)
            {
                //가격 조정 후 원가계산서 엑셀파일 생성
                FillCostAccount.FillBiddingCosts();

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                saveFileDialog.Filter = "Microsoft Excel (*.xlsx)|*.xlsx";
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = "원가계산서_세부결과";
                saveFileDialog.OverwritePrompt = true;


                if (saveFileDialog.ShowDialog() == true)
                {
                    string file = saveFileDialog.FileName.ToString(); //경로와 파일명 저장
                    string xlsxFolder = Data.work_path;
                    string costFile = Path.Combine(xlsxFolder, "원가계산서_세부결과.xlsx");

                    File.Move(costFile, file);
                    DisplayDialog("저장되었습니다.", "Save");
                }
                else
                {
                    DisplayDialog("취소되었습니다.", "Error");
                }
            }

            // 계산 안되어 있을 경우
            else
            {
                DisplayDialog("계산을 먼저 실행해주세요.", "Error");
            }
        }


        // ------------------------- 입찰 내역 저장 버튼 -------------------------------------------------------------------------------------------------------------------------------------- //
        private void SaveBiddingZipBtnClick(object sender, System.EventArgs e)
        {
            // TargetRate가 계산 되어 있을 경우
            if (isCalculate)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                saveFileDialog.Filter = "Zip 압축 파일 (*.zip)|*.zip";
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = "입찰내역";
                saveFileDialog.OverwritePrompt = true;


                if (saveFileDialog.ShowDialog() == true)
                {
                    string file = saveFileDialog.FileName.ToString(); //경로와 파일명 저장
                    string biddingFolder = Data.work_path; //입찰 내역 경로
                    string biddingZipFile = Path.Combine(biddingFolder, "입찰내역.zip");

                    Directory.Move(biddingZipFile, file);
                    DisplayDialog("저장되었습니다.", "Save");
                }
                else
                {
                    DisplayDialog("취소되었습니다.", "Error");
                }


            }

            // 계산 안되어 있을 경우
            else
            {
                DisplayDialog("계산을 먼저 실행해주세요.", "Error");
            }

        }
    }
}