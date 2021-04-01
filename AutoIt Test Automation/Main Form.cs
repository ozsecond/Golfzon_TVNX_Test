using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using AutoIt;
using Accord.Video.FFMPEG;
using Tesseract;
using OpenCvSharp;
using OpenCvSharp.Extensions;

// TODO 


    


namespace AutoIt_Test_Automation
{
    public partial class FORMmain : Form
    {
        #region ---- 전역변수 ----

        // 어플리케이션 이름 (향후 선택가능 해야 함)
        private string GS_TYPE;
        string DEFAULT_APP_NAME = "GolfZonCilent"; // VISION



        // 클래스 컨트롤 인스턴스
        LogControl LOG = new LogControl();



        // 템프폴더 패스
        private string _imagePath;
        private string _tempPath = Application.StartupPath + @"\temp\";
        private string _resultPath = Application.StartupPath + @"\Result\";



        // 이미지 시퀀스 저장 리스트
        private List<string> _inputImageSequence = new List<string>();
        private int _fileCount = 1;




        // 해상도 변수
        public static int _resX = 1920;
        public static int _resY = 1080;
        public static int _left = 0;
        public static int _top = 0;



        // GS 패스워드
        private string _password;



        // 작업홀         
        private string[] _workHoles;        
        private int _repeatCount = 1;



        // 스레딩타이머 딜리게이트
        delegate void TimerEventFiredDelegate();



        // 키보드 이벤트 처리 API
        [DllImport("user32.dll")]
        public static extern void KeyboardEvent(uint vk, uint scan, uint flags, uint extraInfo);



        // 한글 초중종성 유니코드 분석용 
        public static readonly char[] CHOSUNG_TABLE = {'ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ'};
        public static readonly char[] JUNGSUNG_TABLE = {'ㅏ', 'ㅐ', 'ㅑ', 'ㅒ', 'ㅓ', 'ㅔ', 'ㅕ', 'ㅖ', 'ㅗ', 'ㅘ', 'ㅙ', 'ㅚ', 'ㅛ', 'ㅜ', 'ㅝ', 'ㅞ', 'ㅟ', 'ㅠ', 'ㅡ', 'ㅢ', 'ㅣ'};
        public static readonly char[] JONGSUNG_TABLE = { ' ', 'ㄱ', 'ㄲ', 'ㄳ', 'ㄴ', 'ㄵ', 'ㄶ', 'ㄷ', 'ㄹ', 'ㄺ', 'ㄻ', 'ㄼ', 'ㄽ', 'ㄾ', 'ㄿ', 'ㅀ', 'ㅁ', 'ㅂ', 'ㅄ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ' };
        

        #endregion



        public FORMmain()
        {
            // 나눔명조 폰트 추가
            PrivateFontCollection pfc = new PrivateFontCollection();
            pfc.AddFontFile(Application.StartupPath + @"\Font\NanumMyeongjo.ttf");
            this.Font = new Font(pfc.Families[0], 11f);            

            InitializeComponent();

            // 빌드버전 표시
            LBLbuildVer.Text = "ver." + Assembly.GetExecutingAssembly().GetName().Version;            

            // 로그 초기화            
            LOG.SetLogFileName();            

            // 템프 및 이미지 폴더 생성
            InitTempAndResultFolder();
        }



        // 스타트 버튼 클릭 이벤트
        private void BTNstart_Click(object sender, EventArgs e)
        {
            // 패스워드 변수 입력
            _password = TXTpassword.Text;
            LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : GS 패스워드 : " + _password);

            // GS 타입 반영 << 향후 수정 필요
            GS_TYPE = CBBgsType.Text;

            // GS 타입에 따른 어플리케이션 네임 배정
            switch (GS_TYPE)
            {
                case "VISION":
                    DEFAULT_APP_NAME = "GolfZonCilent";
                    _imagePath = Application.StartupPath + @"\Images\Vision\";
                    break;

                case "TWOVISION":
                    DEFAULT_APP_NAME = "GolfZonClient";
                    _imagePath = Application.StartupPath + @"\Images\Twovision\";
                    _resX = 3840;
                    _resY = 2160;                    

                    break;

                default:
                    DEFAULT_APP_NAME = "GolfZonCilent";
                    _imagePath = Application.StartupPath + @"\Images\Vision\";
                    break;
            }

            // 작업홀을 읽어옴
            string[] tempWorkHoles = TXThole.Text.Split(',');

            Array.Resize(ref _workHoles, tempWorkHoles.Length);            

            for (int i = 0; i < tempWorkHoles.Length; i++)
            {
                _workHoles[i] = tempWorkHoles[i].Trim();
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : _workHole = " + _workHoles[i].ToString());
            }            

            // 반복횟수 읽어옴
            int repeatCount = int.Parse(TXTrepeat.Text);
            LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : repeatCount = " + repeatCount);


            // 매크로 시작 (타이틀 화면에서)
            LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : InitClient() 실행");
            InitClient();

            // 반복횟수 만큼 반복
            for (int i = 0; i < repeatCount; i++)
            {
                // 작업홀 만큼 반복
                for (int j = 0; j < _workHoles.Length; j++)
                {
                    LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : AdCaptureSelectedHoles 실행, repeatCount = " + i.ToString() 
                        + " ,workHoles = " + _workHoles[j].ToString());

                    // _workHoles[j]가 숫자인지 문자열(17-18)인지 판단하여 함수를 구분
                    // 숫자로 변환이 가능할 경우 해당홀+다음홀에 나오는 로직 태움
                    if (int.TryParse(_workHoles[j], out int workHole))
                    {
                        int[] workHoles = new int[1];
                        workHoles[0] = workHole;

                        AdCaptureSelectedHoles(workHoles);
                    }

                    // 숫자로 변환이 불가능할 경우(ex. 1-2, 17-18 등), 연속으로 홀스킵을 시행하는 로직 태움.
                    else
                    {
                        // 잘못된 형식이 들어왔을 경우 오류메시지 드로우
                        try
                        {
                            string throwMessage = MethodBase.GetCurrentMethod().Name + " : 비정상적인 형식의 홀번호";

                            string[] temp = _workHoles[j].Split('-');
                            if (temp.Length != 2)
                            {
                                throw new Exception(throwMessage);
                            }

                            int top;
                            int bottom;
                            int mid;

                            // 둘 중 높은 숫자를 top으로 보냄
                            if (!int.TryParse(temp[0], out bottom))
                            {
                                throw new Exception(throwMessage);
                            }

                            if (!int.TryParse(temp[1], out top))
                            {
                                throw new Exception(throwMessage);
                            }

                            // 순서가 잘못됬을 경우 바로잡는 코드
                            if (top < bottom)
                            {
                                mid = top;
                                top = bottom;
                                bottom = mid;
                            }

                            // 홀 범위를 벗어날 경우 익셉션 드로우
                            if (bottom < 1 || top > 18)
                            {
                                throw new Exception(throwMessage);
                            }

                            // workHoles 배열에 Bottom부터 Top까지 +1 단위로 입력
                            int[] workHoles = new int[top - bottom + 1];
                            workHoles[0] = bottom;
                            
                            for (int x = 1; x < top - bottom + 1; x++)
                            {
                                workHoles[x] = workHoles[x - 1] + 1;
                            }                            

                            // 실행
                            AdCaptureSelectedHoles(workHoles);

                        }
                        catch (Exception error)
                        {
                            LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : " + error.Message);
                            Application.ExitThread();
                            Environment.Exit(0);
                        }
                    }

                    LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : AdCaptureSelectedHoles 완료, repeatCount = " + i.ToString()
                        + " ,workHoles = " + _workHoles[j].ToString());
                }

                _repeatCount++;
            }


            Thread.Sleep(3000);
            LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 코드 실행 완료");

            AutoItX.WinActivate(Application.ProductName);
            MessageBox.Show("실행 완료!!!");
        }



        // 게스트 로그인 후 모드선택까지 이동
        private void InitClient()
        {
            try
            {
                // 창 활성화
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 어플리케이션 창 활성화 시도");
                AutoItX.WinActivate(DEFAULT_APP_NAME);


                // 어플리케이션 핸들을 받아오지 못하면 강제 종료
                string handle = AutoItX.WinGetHandleAsText(DEFAULT_APP_NAME);
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : handle = " + handle);

                if (handle == "0x00000000")
                {
                    LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 어플리케이션 창 핸들을 받아오지 못함");
                    throw new Exception("어플리케이션 창 핸들을 받아오지 못함");
                }
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 어플리케이션 창 활성화 성공");


                // 창 활성화까지 대기
                Thread.Sleep(3000);

                // 아무키나 입력해서 비번 창을 뛰움
                AutoItX.Send("{ENTER}");

                // 비번 창 활성화까지 대기
                Thread.Sleep(2000);

                // 비번 창 텍스트 박스 클릭 (혹시 포커스가 안가있을지도 모를 사태 대비)
                AutoItX.MouseClick("LEFT", 881, 436);
                Thread.Sleep(3000);

                // 패스워드 입력
                AutoItX.Send(_password);
                Thread.Sleep(1000);
                AutoItX.Send("{Enter}");
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + "패스워드 입력 완료");

                // 플레이어 설정 전환까지 대기
                AutoItX.MouseMove(_resX, _resY);
                Thread.Sleep(3000);

                // 정상적으로 플레이어 설정 진입했는지 확인절차                
                Bitmap playerConfig = CapturePartialScreen(317, 289, 169, 38, _tempPath + "플레이어설정확인");
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 플레이어 설정 진입 확인");

                if (ImageToString(playerConfig).Trim() != "플레이어 설정")
                {
                    LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 플레이어 설정이 아님. 매크로 종료");
                    throw new Exception("플레이어 설정 진입 실패");
                }

                // 게스트 등록 버튼 클릭
                AutoItX.MouseClick("LEFT", 890, 225);
                Thread.Sleep(2000);

                // 다음 버튼 클릭
                AutoItX.MouseClick("LEFT", 1477, 924);

                // 모드선택 - 대회정보를 수신할때까지 넉넉히 대기 
                LOG.WriteLog(TXTlog, "모드선택 진입");
                Thread.Sleep(5000);
            }
            catch (Exception error)
            {
                AutoItX.WinActivate(Application.ProductName);
                MessageBox.Show(error.Message);

                Application.ExitThread();
                Environment.Exit(0);
            }
        }



        // 인게임 내 홀스킵 이동 광고 캡쳐
        private void AdCaptureSelectedHoles(int[] holes)
        {
            try
            {
                // 홀은 시작홀
                int firstHole = holes[0];

                // 라이브 광고 팝업 예외처리
                AutoItX.MouseClick("LEFT", 1013, 888);
                Thread.Sleep(3000);

                // 모드선택 진입확인
                IsModeSelect();

                // 스트로크 선택
                AutoItX.MouseClick("LEFT", 739, 664);
                Thread.Sleep(3000);

                // 스트로크 상태에서 다음
                AutoItX.MouseClick("LEFT", 1499, 914);

                // CC 선택 - 넉넉히 대기
                LOG.WriteLog(TXTlog, "CC선택 진입");
                Thread.Sleep(5000);

                // CC 입력 텍스트박스 선택
                AutoItX.MouseClick("LEFT", 1469, 177);
                Thread.Sleep(3000);
                InputGameText(TXTccName.Text, true);
                Thread.Sleep(3000);
                AutoItX.Send("{Enter}");

                // CC 검색시간 대기
                Thread.Sleep(3000);

                // CC 선택에서 다음
                AutoItX.MouseClick("LEFT", 1499, 914);
                Thread.Sleep(3000);

                // 코스매니저와 볼설정 팝업창
                AutoItX.MouseClick("LEFT", 838, 783);

                // 라운드 설정으로 이동
                LOG.WriteLog(TXTlog, "라운드설정 진입");
                Thread.Sleep(3000);

                // 기타설정 클릭
                AutoItX.MouseClick("LEFT", 1594, 487);
                Thread.Sleep(2000);

                // 시작홀 선택                
                ClickStartHole(firstHole);

                // 라운드시작 클릭 
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 라운드 시작 클릭");
                AutoItX.MouseClick("LEFT", 1499, 914);
                Thread.Sleep(2000);

                // 서비스 이용 상세내역 팝업창
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name +  " : 서비스 이용 상세내역 팝업창 진입");
                AutoItX.MouseClick("LEFT", 833, 580);
                Thread.Sleep(1000);

                // 기본 짧은 광고로딩 녹화 시간 (10초)
                int recordingTime = 10000;

                // 에티켓 광고 포함 (+10초)
                recordingTime += 10000;

                // 1홀 로딩 - 1홀 이나 10홀일경우 긴 로딩광고 녹화 필요. (+ 25초)
                if (firstHole == 1 || firstHole == 10)
                {
                    recordingTime += 25000;
                }

                string videoFileName = firstHole.ToString() + "H 로딩광고_" + _repeatCount;

                // 비디오 레코딩 시작
                RecordVideo(recordingTime, videoFileName);

                // 녹화가 끝난 후 여유있게 10초 대기
                Thread.Sleep(10000);

                // 진입 후 원하는 홀에 들어왔는지 확인                
                HoleImageCheck(firstHole);

                // 홀 진입 상태                
                Thread.Sleep(3000);


                foreach (int hole in holes)
                {
                    LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : " + hole.ToString() + "H 진입완료");

                    // 나스모 광고 촬영
                    if (CHKnasmo.Checked)
                    {
                        AutoItX.Send("{F8}");
                        Thread.Sleep(3000);
                        CaptureFullScreen(_resultPath + hole.ToString() + "H 나스모 광고_" + _repeatCount);
                        AutoItX.Send("{F8}");
                        Thread.Sleep(3000);
                    }

                    // 18홀일 경우 이후의 액션을 스킵하고 바로 종료
                    if (hole != 18)
                    {
                        // 홀스킵 사용
                        AutoItX.Send("{F5}");
                        Thread.Sleep(2000);

                        // 홀스킵 팝업창 진입
                        LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 홀스킵 팝업창 진입");
                        AutoItX.Send("{ENTER}");
                        Thread.Sleep(2000);

                        // 스코어보드 광고 촬영
                        CaptureFullScreen(_resultPath + hole.ToString() + "H 스코어보드 광고_" + _repeatCount);

                        // 다음홀 진입 전 짧은 로딩광고
                        recordingTime = 10000;

                        // 만약 다음홀이 10홀일 경우 +25초 추가
                        if (hole + 1 == 10)
                        {
                            recordingTime += 25000;
                        }

                        videoFileName = (hole + 1).ToString() + "H 로딩광고_" + _repeatCount;

                        // 비디오 레코딩 시작
                        RecordVideo(recordingTime, videoFileName);

                        Thread.Sleep(10000);

                        // 진입 후 원하는 홀에 들어왔는지 확인                
                        HoleImageCheck(hole + 1);

                        Thread.Sleep(3000);
                    }
                }

                // 종료로직 수행
                IngameQuit();


                // 종료 로딩


                // 모드선택으로 돌아올때까지 대기
                Thread.Sleep(15000);


                // 홀/시간 설정 진입
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 모드선택 홀/시간설정 팝업 진입확인");

                Bitmap holeSetting = CapturePartialScreen(607, 339, 706, 77, _tempPath + "홀시간설정");
                Bitmap holeSetting2 = new Bitmap(_imagePath + "LobbyHoleSettingPop.png");
                if (CompareImages(holeSetting, holeSetting2) < 98)
                {
                    throw new Exception("홀/시간설정 팝업 진입 실패");
                }

                AutoItX.MouseClick("LEFT", 1016, 440);
                Thread.Sleep(3000);
                AutoItX.Send(_password);
                Thread.Sleep(3000);
                AutoItX.Send("{ENTER}");
                Thread.Sleep(3000);
            }

            catch(Exception error)
            {
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : " + error.Message);
                AutoItX.WinActivate(Application.ProductName);
                MessageBox.Show(error.Message);

                Application.ExitThread();
                Environment.Exit(0);
            }
        }



        // 모드선택 확인 절차
        private void IsModeSelect()
        {
            // 모드선택 진입 확인            
            Bitmap modeSelect = CapturePartialScreen(324, 396, 107, 33, _tempPath + "모드선택확인");
            LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 모드선택 진입 확인");

            // 모드선택 -> 모드선랙으로 읽음 *TS 엔진 한계
            if (ImageToString(modeSelect).Trim() != "모드선랙")
            {
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + "모드선택 진입 실패. 매크로 종료");
                throw new Exception("모드선택 진입 실패");
            }            
        }



        // 인게임 종료 팝업창 클릭 버그 해결
        private void IngameQuit()
        {
            try
            {
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 게임종료 로직 수행");

                // 게임종료
                AutoItX.Send("{ESC}");
                Thread.Sleep(3000);


                // 관리자 비번 진입
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 게임종료 관리자비번 진입");
                AutoItX.Send(_password);
                Thread.Sleep(3000);
                AutoItX.Send("{ENTER}");
                Thread.Sleep(3000);


                // 유지종료 (한번하면 잘 안됨)
                int tolerance = 100;
                int timeOut = 0;

                // 종료팝업이 없어질때까지 무한 반복
                while (tolerance >= 98)
                {
                    Bitmap screenShot = CapturePartialScreen(606, 340, 708, 178, _tempPath + "게임종료팝업");
                    Bitmap img = new Bitmap(_imagePath + "IngameQuitPop.png");
                    tolerance = CompareImages(screenShot, img);
                    timeOut++;

                    AutoItX.MouseClick("LEFT", 961, 636);
                    LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 유지종료 클릭 = " + timeOut.ToString());

                    Thread.Sleep(1000);
                }
            }
            catch (Exception error)
            {
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : " + error.Message);
                AutoItX.WinActivate(Application.ProductName);
                MessageBox.Show(error.Message);

                Application.ExitThread();
                Environment.Exit(0);
            }
        }
        


        // 시작홀을 받아서 라운드설정-홀선택에서 마우스 좌표를 찍게 해줌
        private void ClickStartHole (int startHole)
        {
            LOG.WriteLog(TXTlog, "시작홀 : " + startHole.ToString());
            
            // 칼럼번호 1홀 = 1
            int column = startHole;

            // 1홀 스타트 클릭 위치
            int clickX = 950;
            int clickY = 600;

            // 스타트 홀이 아랫줄일 경우 Y좌표 이동
            if (startHole > 9)
            {
                clickY = 685;
                column -= 9;
            }

            // 칼럼번호 만큼 x클릭 위치를 우측으로 이동시킴
            for (int i = 1; i < column; i++)
            {
                clickX += 60;
            }

            // 좌표 클릭
            AutoItX.MouseClick("LEFT", clickX, clickY);            
            Thread.Sleep(2000);
        }


        
        // 인게임 홀 번호 대조
        private int GetHoleNumber()
        {
            Bitmap img = CapturePartialScreen(75, 18, 47, 53, _tempPath + "홀번호");
            List<string> fileNames = new List<string> ();
            int holeNo = 0;

            // 비전의 경우 VisionHoleImg 폴더 사용 ★★★ 다른 GS에서는 다른 폴더를 이용해야 함
            DirectoryInfo di = new DirectoryInfo(_imagePath);
            
            foreach (FileInfo file in di.GetFiles())
            {
                Bitmap holeImg = new Bitmap(file.FullName);

                // 이미지 정합성이 98% 이상일경우 해당 홀로 판정
                LOG.WriteLog(TXTlog, "::GetHoleNumber : " + "홀번호.png" + " vs " + file.Name);
                if (CompareImages(img, holeImg) >= 98)
                {
                    // 홀이미지 파일명에서 확장자 앞의 두글자만 가져옴 
                    string temp = file.FullName.Substring(file.FullName.Length - 6, 2);                    

                    // 첫자리가 _일 경우 1자리수로 판별하고 뒷자리수를 가져옴
                    if (temp.Contains("_"))
                    {
                        temp = temp.Substring(1, 1);
                    }
                    holeNo = int.Parse(temp);

                    return holeNo;
                }
            }

            return holeNo;
        }



        // 홀 대기 
        private void HoleImageCheck(int hole)
        {
            int currentHole = GetHoleNumber();
            int overTime = 0;

            // 홀 이미지가 매칭될때까지 1초 단위로 확인
            while (hole != currentHole)
            {
                LOG.WriteLog(TXTlog, "::HoleImageCheck : 홀 이미지 매칭 중 : currentHole = " + currentHole.ToString());

                Thread.Sleep(1000);
                overTime += 1;
                currentHole = GetHoleNumber();

                // 만약 120초간 대기에도 매칭이 되지 않으면 강제 종료
                if (overTime > 120)
                {
                    LOG.WriteLog(TXTlog, "::HoleImageCheck : 매크로 종료 : 홀 이미지 매칭 대기 초과");
                    throw new Exception("::HoleImageCheck : 홀 이미지 매칭 대기 초과");
                }
            }

            return;
        }
        


        // 범용 이미지 체크 로직 (체크에 실패할 경우 어플리케이션 강제 종료
        private void CommonImageCheck(int x, int y, int width, int height, string filePath, int tolerance)
        {
            Bitmap screenShot = CapturePartialScreen(x, y, width, height, _tempPath + "imgChk1");            
            Bitmap img = new Bitmap(filePath);

            LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : imgChk1 vs " + filePath + " , tolerance = " + tolerance.ToString());

            // 이미지가 톨레랑스보다 낮을 경우 어플 종료
            if (CompareImages(screenShot, img) < tolerance)
            {
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 이미지 체크 실패");
                Application.ExitThread();
                Environment.Exit(0);
            }
        }

        
        
        // 한글키를 눌러주는 키보드 이벤트
        private void PressHangulKey()
        {
            KeyboardEvent((byte)Keys.HanguelMode, 0, 0x00, 0);
            KeyboardEvent((byte)Keys.HanguelMode, 0, 0x02, 0);
        }



        // 텍스트 입력 시 게임엔진에 맞게 입력
        private void InputGameText(string text, bool isBeforeCharKor)
        {
            // isBeforeCharKor 이 0 일경우 영어시작으로 인식
            // isBeforeCharKor 이 1 일경우 한글시작으로 인식


            // 1글자 단위로 분리하여 한자씩 전송
            foreach (char ch in text)
            {
                string sendText = "";
                
                switch (CheckString(ch))
                {
                    case "K": 
                        // 이전에 입력된 글자가 한글이 아니면 키 한영전환 
                        if (!isBeforeCharKor)
                        {
                            PressHangulKey();
                            isBeforeCharKor = true;
                        }

                        sendText = ch.ToString();
                        
                        break;

                    case "E":
                        // 이전에 입력된 글자가 한글이었으면 키 한영전환 
                        if (isBeforeCharKor)
                        {
                            PressHangulKey();
                            isBeforeCharKor = false;
                        }

                        sendText = ch.ToString();
                        
                        break;

                    case "S":
                        sendText = "{SPACE}";
                        break;

                        // 숫자나 특문은 처리 필요없음
                    default:
                        sendText = ch.ToString();
                        break;
                }

                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 키보드 입력 신호 전달 = " + sendText);
                AutoItX.Send(sendText);

                // 정상적인 입력을 위해 유휴시간을 둠
                Thread.Sleep(200);
            }
        }



        // 입력된 텍스트가 영어인지 한글인지 자음단위로 판별
        private string CheckString(char ch)
        {
            // 한글
            if ((0xAC00 <= ch && ch <= 0xD7A3) || (0x3131 <= ch && ch <= 0x318E))
            {
                return "K";
            }

            // 영어
            else if ((0x61 <= ch && ch <= 0x7A) || (0x41 <= ch && ch <= 0x5A))
            {
                return "E";
            }

            // 숫자
            else if (0x30 <= ch && ch <= 0x39)  
            {
                return "N";
            }

            // 빈칸
            else if (ch == ' ')
            {
                return "S";
            }

            return "D";
        }
        


        // 한글의 초중종성 분해 -> 투비전 화상키보드 인식용
        private void TouchHangulKeyBoard(char ch)
        {
            // 한글이 아니면 null을 리턴
            if (CheckString(ch) != "K")
            {
                return;
            }

            int temp = Convert.ToUInt16(ch);
            int nUniCode = temp - 0xAC00;

            // 초성 중성 종성 변수
            int cho = nUniCode / (21 * 28);
            nUniCode = nUniCode % (21 * 28);

            int jung = nUniCode / 28;
            int jong = nUniCode % 28;

            // 마우스 클릭용 좌표
            int x = 0;
            int y = 0;
            bool isNeedShift = false;

            char[] jaso = new char[3];
            jaso[0] = CHOSUNG_TABLE[cho];
            jaso[1] = JUNGSUNG_TABLE[jung];
            jaso[2] = JONGSUNG_TABLE[jong];

            for (int i = 0; i < jaso.Length; i++)
            {
                // 화상키보드 좌표와 연동해 클릭
                switch (jaso[i])
                {
                    case 'ㄱ':
                        x = 774;
                        y = 692;
                        break;

                    case 'ㄲ':
                        isNeedShift = true;
                        x = 774;
                        y = 692;
                        break;

                    case 'ㄴ':
                        x = 645;
                        y = 770;
                        break;

                    case 'ㄷ':
                        x = 687;
                        y = 704;
                        break;

                    case 'ㄸ':
                        isNeedShift = true;
                        x = 687;
                        y = 704;
                        break;

                    case 'ㄹ':
                        x = 820;
                        y = 773;
                        break;

                    case 'ㅁ':
                        x = 591;
                        y = 779;
                        break;

                    case 'ㅂ':
                        x = 545;
                        y = 689;
                        break;

                    case 'ㅃ':
                        isNeedShift = true;
                        x = 545;
                        y = 689;
                        break;

                    case 'ㅅ':
                        x = 861;
                        y = 692;
                        break;

                    case 'ㅆ':
                        isNeedShift = true;
                        x = 861;
                        y = 692;
                        break;

                    case 'ㅇ':
                        x = 727;
                        y = 767;
                        break;
    
                    case 'ㅈ':
                        AutoItX.MouseClick("LEFT", 609, 686
    

            case 'ㅉ':
                        AutoItX.MouseClick("LEFT", 459, 857
        

                        Sleep 200
        

                        AutoItX.MouseClick("LEFT", 609, 686
        

                        Sleep 200
        

                        AutoItX.MouseClick("LEFT", 459, 857
    

            case 'ㅊ':
                        AutoItX.MouseClick("LEFT", 750, 846
    

            case 'ㅋ':
                        AutoItX.MouseClick("LEFT", 580, 852
    

            case 'ㅌ':
                        AutoItX.MouseClick("LEFT", 664, 855
    

            case 'ㅍ':
                        AutoItX.MouseClick("LEFT", 825, 851
    

            case 'ㅎ':
                        AutoItX.MouseClick("LEFT", 889, 767
                }

                // 쌍자음일 경우 시프트 클릭
                if (isNeedShift)
                {
                    // 시프트키 클릭
                    AutoItX.MouseClick("LEFT", 459, 857);
                    Thread.Sleep(500);

                    // 자음 클릭
                    AutoItX.MouseClick("LEFT", x, y);
                    Thread.Sleep(500);

                    // 시프트키 끄기
                    AutoItX.MouseClick("LEFT", 459, 857);

                }
                else
                {
                    AutoItX.MouseClick("LEFT", x, y);
                }
            }

            
            



            LOG.WriteLog(TXTlog, "초성 : " + CHOSUNG_TABLE[cho].ToString());
            LOG.WriteLog(TXTlog, "중성 : " + JUNGSUNG_TABLE[jung].ToString());
            LOG.WriteLog(TXTlog, "종성 : " + JONGSUNG_TABLE[jong].ToString());            
        }

        // 자소를 받아 터치스크린을 클릭
        private void TouchJamo(int x, int y)
        {
            AutoItX.MouseClick("LEFT", 459, 857);
            Thread.Sleep(500)
        }




        // 템프파일 삭제
        private void DeleteTempFiles()
        {
            DirectoryInfo di = new DirectoryInfo(_tempPath);

            foreach (FileInfo fi in di.GetFiles())
            {
                fi.Delete();
            }
        }



        #region ---- 사용하지 않는 스레드 기능 ----

        // 스레딩 타이머
        private void VRCallBack(object status)
        {
            BeginInvoke(new TimerEventFiredDelegate(ThreadVideoRecording));
        }

        private void ThreadVideoRecording()
        {
            using (Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.CopyFromScreen(0, 0, 0, 0, bitmap.Size);
                }

                string fileName = _tempPath + "videoTemp_" + _fileCount + ".png";
                bitmap.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                _inputImageSequence.Add(fileName);
                _fileCount++;

                bitmap.Dispose();
            }
        }

        // 딜레이
        private void Delay(int time)
        {
            DateTime ThisMoment = DateTime.Now;
            TimeSpan duration = new TimeSpan(0, 0, 0, 0, time);
            DateTime AfterWards = ThisMoment.Add(duration);

            while (AfterWards >= ThisMoment)
            {
                Application.DoEvents();
                ThisMoment = DateTime.Now;
                //LOG.WriteLog(TXTlog, "::Delay Afterward = " + AfterWards.ToString() + " , ThisMoment = " + ThisMoment.ToString());
            }
        }

        #endregion



        #region ---- 캡쳐 메소드들 ----

        // 비디오 레코드
        private void RecordVideo(int term, string fileName)
        {
            LOG.WriteLog(TXTlog, "비디오 레코드 시작 : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            
            DateTime targetTime = DateTime.Now.AddMilliseconds(term);

            // 초당 8 ~ 20프레임이 한계
            while (DateTime.Now <= targetTime)
            {
                Recording();
            }


            // 생성된 이미지 파일 수로 영상의 fps를 계산
            double fps = (double)_fileCount / ((double)term / 1000);
            fps = Math.Round(fps);

            LOG.WriteLog(TXTlog, "::RecordVideo : _fileCount = " + _fileCount.ToString("N0") 
                + " ,term = " + term.ToString("N0") 
                + "ms ,fps = " + fps.ToString());

            if (fps != 0)
            {
                SaveVideo((int)fps, fileName);
            }
            else
            {
                LOG.WriteLog(TXTlog, "계산된 fps가 0이어서 영상을 인코딩할 수 없음.");
            }
                
        }



        // 비디오 녹화 <- 타이머 틱으로 처리
        private void Recording()
        {
            using (Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.CopyFromScreen(0, 0, 0, 0, bitmap.Size);
                }

                string fileName = _tempPath+ "videoTemp_" + _fileCount + ".png";
                bitmap.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                _inputImageSequence.Add(fileName);
                _fileCount++;

                bitmap.Dispose();
            }
        }
        


        // 비디오 파일 저장
        private void SaveVideo(int fps, string fileName)
        {
            try
            {
                using (VideoFileWriter vfw = new VideoFileWriter())
                {
                    vfw.Open(_resultPath + fileName + ".mp4",
                        Screen.PrimaryScreen.Bounds.Width,
                        Screen.PrimaryScreen.Bounds.Height,
                        fps,
                        VideoCodec.MPEG4
                        );
                    
                    LOG.WriteLog(TXTlog, "::SaveVideo : 비디오 인코딩 시작 : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    LOG.WriteLog(TXTlog, "::SaveVideo : 파일 카운트 : " + _fileCount.ToString("N0"));
                    
                    // 이미지 시퀀스 비디오 인코딩
                    foreach (string imageLoc in _inputImageSequence)
                    {
                        Bitmap imageFrame = Image.FromFile(imageLoc) as Bitmap;
                        vfw.WriteVideoFrame(imageFrame);
                        imageFrame.Dispose();
                    }

                    LOG.WriteLog(TXTlog, "::SaveVideo : 비디오 인코딩 완료 : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    // 파일카운트, 저장변수 초기화
                    _fileCount = 1;
                    _inputImageSequence.Clear();

                    // 템프폴더 삭제
                    DeleteTempFiles();

                    vfw.Close();
                }

                LOG.WriteLog(TXTlog, "::SaveVideo : 비디오파일 저장 완료 : " + _resultPath + fileName + ".mp4");
            }
            catch (Exception e)
            {
                LOG.WriteLog(TXTlog, e.Message);
            }
            
        }



        // 전체 스크린 캡쳐
        private void CaptureFullScreen(string fileName)
        {
            LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 스크린 샷 캡쳐 시도 (" + fileName + ") _resX : " + _resX.ToString());
            
            try
            {
                Bitmap screenshot = new Bitmap(_resX, _resY, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                Graphics gr = Graphics.FromImage(screenshot);
                gr.CopyFromScreen(_left, _top, 0, 0, screenshot.Size);
                screenshot.Save(fileName + ".png", System.Drawing.Imaging.ImageFormat.Png);
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 스크린 샷 캡쳐 성공 (" + fileName + ")");
            }
            catch (Exception error)
            {
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 스크린 샷 캡쳐 실패 : " + error.Message);
            }
        }



        // 부분 스크린 캡쳐
        private Bitmap CapturePartialScreen(int x, int y, int width, int height, string fileName)
        {
            LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 스크린 샷 캡쳐 시도 (" + fileName + ")");

            try
            {
                Bitmap screenshot = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                Graphics gr = Graphics.FromImage(screenshot);
                gr.CopyFromScreen(x, y, 0, 0, screenshot.Size);
                screenshot.Save(fileName + ".png", System.Drawing.Imaging.ImageFormat.Png);
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 스크린 샷 캡쳐 성공 (" + fileName + ")");
                return screenshot;
            }
            catch (Exception error)
            {
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 스크린 샷 캡쳐 실패 : " + error.Message);
            }

            return null;
        }
        #endregion



        // 테스트용 버튼 1 클릭 이벤트
        private void button1_Click(object sender, EventArgs e)
        {
            _resX = 3840;
            _resY = 1080;
            _left = -1920;
            DEFAULT_APP_NAME = "Golfzon";


            try
            {
                // 창 활성화            
                AutoItX.WinActivate(DEFAULT_APP_NAME);
                Thread.Sleep(1000);

                // 투비전은 터치모니터 스크린이 메인이기 때문에 포커스를 여기로 맞춰야함
                AutoItX.WinActivate("GolfzonTouchMonitor");

                // 어플리케이션 핸들을 받아오지 못하면 강제 종료
                string handle = AutoItX.WinGetHandleAsText("GolfzonTouchMonitor");
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : handle = " + handle);

                if (handle == "0x00000000")
                {
                    LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 어플리케이션 창 핸들을 받아오지 못함");
                    throw new Exception("어플리케이션 창 핸들을 받아오지 못함");
                }
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : 어플리케이션 창 활성화 성공");


                Thread.Sleep(3000);

                // 비밀번호 텍스트박스 클릭
                AutoItX.MouseClick("LEFT", 998, 441);
                Thread.Sleep(3000);

                // 비번 입력
                InputGameText(TXTpassword.Text, false);
                Thread.Sleep(1000);

                AutoItX.Send("{ENTER}");
                Thread.Sleep(3000);
                AutoItX.Send("{ENTER}");
                Thread.Sleep(3000);

                // 플레이어 설정 이동               

                MessageBox.Show("버튼 1 실행 완료");
            }
            catch(Exception error)
            {
                LOG.WriteLog(TXTlog, MethodBase.GetCurrentMethod().Name + " : " + error.Message);
                AutoItX.WinActivate(Application.ProductName);
                MessageBox.Show(error.Message);

                Application.ExitThread();
                Environment.Exit(0);
            }

        }



        private void button2_Click(object sender, EventArgs e)
        {
            string ccName = TXTccName.Text;
            foreach(char ch in ccName)
            {
                TouchHangulKeyBoard(ch);
            }
        }
        


        private void button3_Click(object sender, EventArgs e)
        {
            
        }



        #region ---- 이미지 처리 메소드들 ----

        // OpenCV 이진화 이미지 전처리
        private Bitmap ImageBinary(Bitmap img)
        {
            try
            {
                // 비트맵을 MAT로 변환
                Mat src = BitmapConverter.ToMat(img);
                Mat gray = new Mat();
                Mat bin = new Mat();

                // 그레이스케일로 변환
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);
                LOG.WriteLog(TXTlog, "이미지 그레이스케일 변환 완료");

                // 이진화
                Cv2.Threshold(gray, bin, 0, 255, ThresholdTypes.Otsu);


                // 이진화 파일 저장
                Cv2.ImWrite(_tempPath + "grayscale.png", gray);
                Cv2.ImWrite(_tempPath + "bin.png", bin);

                LOG.WriteLog(TXTlog, "이미지 이진화 변환 완료");
                return BitmapConverter.ToBitmap(bin);
            }
            catch (Exception error)
            {
                LOG.WriteLog(TXTlog, error.Message);
            }

            return null;
        }


        
        // 이미지 확대
        private Bitmap ImageUpScale(Bitmap img)
        {
            Bitmap upScale;

            Mat src = BitmapConverter.ToMat(img);
            Mat dst = new Mat();

            Cv2.PyrUp(src, dst);
            Cv2.ImWrite(_tempPath + "upscale.png", dst);

            upScale = BitmapConverter.ToBitmap(dst);
            LOG.WriteLog(TXTlog, "이미지 업스케일링 변환 완료");

            return upScale;
        }

        

        // 이미지 축소
        private Bitmap ImageDownScale(Bitmap img)
        {
            Bitmap downScale;

            Mat src = BitmapConverter.ToMat(img);
            Mat dst = new Mat();

            Cv2.PyrDown(src, dst);
            Cv2.ImWrite(_tempPath + @"\downScale.png", dst);

            downScale = BitmapConverter.ToBitmap(dst);
            LOG.WriteLog(TXTlog, "이미지 다운스케일링 변환 완료");

            return downScale;
        }


        
        // 이미지 역상(색상 반전)
        private Bitmap ImageReverse(Bitmap img)
        {
            Mat src = BitmapConverter.ToMat(img);
            Mat dst = new Mat();

            Cv2.BitwiseNot(src, dst);
            Cv2.ImWrite(_tempPath + "reverse.png", dst);

            LOG.WriteLog(TXTlog, "이미지 역상 변환 완료");

            return BitmapConverter.ToBitmap(dst);
        }
        


        // 이미지 파일 비트맵 픽셀 비교
        private int CompareImages(Bitmap img1, Bitmap img2)
        {
            if (img1.Width == img2.Width && img1.Height == img2.Height)
            {
                double totalPixel = img1.Width * img1.Height;
                double falseCnt = 0;

                for (int x = 0; x < img1.Width; x++)
                {
                    for (int y = 0; y < img1.Height; y++)
                    {
                        if (img1.GetPixel(x, y) != img2.GetPixel(x, y))
                        {
                            falseCnt++;
                        }
                    }
                }

                double result = 100 - (falseCnt / totalPixel * 100);
                LOG.WriteLog(TXTlog, "::CompareImages : 이미지 비트맵 비교 결과 : " + result.ToString("N2") + "% 일치");

                return (int)result;
            }

            LOG.WriteLog(TXTlog, "::CompareImages : 이미지 비트맵 비교 결과 : 이미지 사이즈가 맞지 않음");

            return -1;
        }



        // OCR로 이미지 파일을 받아 문자로 변환
        private string ImageToString(Bitmap img)
        {
            try
            {
                using (var engine = new TesseractEngine(@"./tessdata", "kor", EngineMode.TesseractOnly))
                {
                    using (var page = engine.Process(img))
                    {
                        LOG.WriteLog(TXTlog, "이미지 Tesseract 결과 : " + page.GetText());
                        return page.GetText();
                    }
                }
            }
            catch (Exception error)
            {
                LOG.WriteLog(TXTlog, error.Message);
            }

            return null;
        }
        #endregion



        // Temp 폴더 생성
        private void InitTempAndResultFolder()
        {
            DirectoryInfo di = new DirectoryInfo(_tempPath);
            if (!di.Exists)
            {
                di.Create();
            }

            DirectoryInfo di2 = new DirectoryInfo(_resultPath);
            if (!di2.Exists)
            {
                di2.Create();
            }
        }               
    }
}

