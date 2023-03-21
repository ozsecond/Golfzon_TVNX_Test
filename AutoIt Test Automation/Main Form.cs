﻿using System;
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
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using AutoIt;
using Accord.Video.FFMPEG;
using Tesseract;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

// TODO 





namespace Golfzon_TVNX_Test
{
    public partial class FORMmain : Form
    {
        #region ---- 전역변수 ----

        // 어플리케이션 이름 (향후 선택가능 해야 함)
        private string GS_TYPE;
        string DEFAULT_APP_NAME = "GolfZonCilent"; // VISION


        // 클래스 컨트롤 인스턴스
        public static FORMmain mainForm;
        LogControl LOG = new LogControl();


        // 템프폴더 패스
        private static string _imagePath = Application.StartupPath + @"\TVNX\";
        private static string _tempPath = Application.StartupPath + @"\Temp\";
        private static string _resultPath = Application.StartupPath + @"\Result\";


        // 이미지 시퀀스 저장 리스트
        private static List<string> _inputImageSequence = new List<string>();
        private static int _fileCount = 1;


        // 해상도 변수
        public static int _resX = 5760;
        public static int _resY = 1080;
        public static int _left = -1920;
        public static int _top = 0;


        // GS 패스워드
        private string _password;


        // 작업홀         
        private string[] _workHoles;
        private int _repeatCount = 1;


        // 키보드 이벤트 처리 API
        [DllImport("user32.dll")]
        public static extern void keybd_event(uint vk, uint scan, uint flags, uint extraInfo);


        // 마우스 이벤트 처리 API
        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
        const uint MOUSEMOVE = 0x0001;      // 마우스 이동
        const uint ABSOLUTEMOVE = 0x8000;   // 전역 위치
        const uint LBUTTONDOWN = 0x0002;    // 왼쪽 마우스 버튼 눌림
        const uint LBUTTONUP = 0x0004;      // 왼쪽 마우스 버튼 떼어짐
        const uint RBUTTONDOWN = 0x0008;    // 오른쪽 마우스 버튼 눌림
        const uint RBUTTONUP = 0x00010;      // 오른쪽 마우스 버튼 떼어짐



        // 한글 초중종성 유니코드 분석용 
        public static readonly char[] CHOSUNG_TABLE = { 'ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ' };
        public static readonly char[] JUNGSUNG_TABLE = { 'ㅏ', 'ㅐ', 'ㅑ', 'ㅒ', 'ㅓ', 'ㅔ', 'ㅕ', 'ㅖ', 'ㅗ', 'ㅘ', 'ㅙ', 'ㅚ', 'ㅛ', 'ㅜ', 'ㅝ', 'ㅞ', 'ㅟ', 'ㅠ', 'ㅡ', 'ㅢ', 'ㅣ' };
        public static readonly char[] JONGSUNG_TABLE = { ' ', 'ㄱ', 'ㄲ', 'ㄳ', 'ㄴ', 'ㄵ', 'ㄶ', 'ㄷ', 'ㄹ', 'ㄺ', 'ㄻ', 'ㄼ', 'ㄽ', 'ㄾ', 'ㄿ', 'ㅀ', 'ㅁ', 'ㅂ', 'ㅄ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ' };


        // 레코딩 스레드 변수
        private Thread _tRecording;
        private static DateTime _recordingStartTime;


        // CC정보 테이블
        DataTable _DTresult;
        DataTable _DTtempResult;
        string _ccName;        
        ArrayList _courseName;


        int[] _par = new int[18];
        int[] _handi = new int[18];
        int[,] _distance = new int[5, 18];

        #endregion


        public FORMmain()
        {
            InitializeComponent();
            mainForm = this;
            string version = "ver." + Assembly.GetExecutingAssembly().GetName().Version;

            // 빌드버전 표시
            LBLbuildVer.Text = version;


            // 로그 초기화            
            LOG.SetLogFileName();
            LOG.Write("Application" + version);


            // 템프 및 이미지 폴더 생성
            InitTempAndResultFolder();


            // 결과 데이터테이블 초기화
            InitResultTable();
        }


        // 결과 데이터테이블 초기화
        private void InitResultTable()
        {
            _DTresult = new DataTable();
            _DTresult.Columns.Add(new DataColumn("CCID", typeof(string)));
            _DTresult.Columns.Add(new DataColumn("CCName", typeof(string)));
            _DTresult.Columns.Add(new DataColumn("CourseName", typeof(string)));
            _DTresult.Columns.Add(new DataColumn("Hole", typeof(int)));
            _DTresult.Columns.Add(new DataColumn("Handicap", typeof(int)));
            _DTresult.Columns.Add(new DataColumn("WebPar", typeof(int)));
            _DTresult.Columns.Add(new DataColumn("CliPar", typeof(int)));
            _DTresult.Columns.Add(new DataColumn("Tee", typeof(string)));
            _DTresult.Columns.Add(new DataColumn("WebDistance", typeof(float)));
            _DTresult.Columns.Add(new DataColumn("CliDistance", typeof(float)));
            _DTresult.Columns.Add(new DataColumn("Result", typeof(string)));


            _DTtempResult = _DTresult.Clone();            
        }


        // 닷컴 CC 데이터 조회 후 결과 테이블 생성
        private void GetWebData(string courseNo)
        {
            // 이전 템프 결과 삭제
            _DTtempResult.Clear();
            

            // 코스번호 입력 (나중엔 csv파일 등에서 읽어와야 함)            
            string url = "http://www.golfzon.com/course/course_detail/R/" + courseNo;


            // 웹 크롤러
            HtmlWeb web = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument htmlDoc = web.Load(url);


            // html 노드에서 정보 추출
            HtmlNode ccNameInfo = htmlDoc.DocumentNode.SelectSingleNode(".//dd[@class='course_posi']/a");
            _ccName = ccNameInfo.InnerText;
            LOG.Write(_ccName);


            _courseName = new ArrayList();
            HtmlNode hallInfo = htmlDoc.DocumentNode.SelectSingleNode(".//div[@class='hall_info']");


            // 코스를 2개 이상 가진 CC는 2개에서 자름
            for (int i = 1; i <= 2; i++)
            {
                float[,] tDis = new float[5, 9];
                int[] par = new int[9];


                // 티 이름 배열
                string[] tee = new string[5];
                tee[0] = "Back";
                tee[1] = "Champ";
                tee[2] = "Front";
                tee[3] = "Senior";
                tee[4] = "Lady";


                HtmlNode courseNameInfo = hallInfo.SelectSingleNode(".//h5[@class='idx_" + i.ToString() + "']/strong");
                if (courseNameInfo != null)
                {
                    _courseName.Add(courseNameInfo.InnerText);
                    LOG.Write("코스추가 = " + courseNameInfo.InnerText);
                }


                HtmlNode parInfo = hallInfo.SelectSingleNode(".//table[@class='idx_" + i.ToString() + "']/thead/tr[2]");
                if (parInfo != null)
                {
                    string tempPar = Regex.Replace(parInfo.InnerText, @"\s", "");
                    for (int h = 0; h < 9; h++)
                    {
                        par[h] = int.Parse(tempPar.Substring(3 + h, 1));
                        LOG.Write("Par Data Added = " + par[h]);
                    }
                }


                HtmlNode teeInfo = hallInfo.SelectSingleNode(".//table[@class='idx_" + i.ToString() + "']/tbody");
                if (teeInfo != null)
                {
                    // 이너텍스트 공백 제거 후 티별 거리 산출
                    string tempTee = Regex.Replace(teeInfo.InnerText, @"\s", "");


                    string[] backT = GetTeeDistance(tempTee, "BackTee", "ChampTee");
                    string[] champT = GetTeeDistance(tempTee, "ChampTee", "FrontTee");
                    string[] frontT = GetTeeDistance(tempTee, "FrontTee", "SeniorTee");
                    string[] seniorT = GetTeeDistance(tempTee, "SeniorTee", "LadyTee");
                    string[] ladyT = GetTeeDistance(tempTee, "LadyTee", "MAP");


                    for (int x = 0; x < 9; x++)
                    {
                        tDis[0, x] = float.Parse(backT[x]);
                        tDis[1, x] = float.Parse(champT[x]);
                        tDis[2, x] = float.Parse(frontT[x]);
                        tDis[3, x] = float.Parse(seniorT[x]);
                        tDis[4, x] = float.Parse(ladyT[x]);
                    }
                }


                // 결과 테이블 작성
                // 홀마다 생성하는 코스이름, 홀, 기준파
                if (courseNameInfo != null)
                {
                    for (int h = 0; h < 9; h++)
                    {
                        DataRow DRh = _DTtempResult.NewRow();
                        DRh["CCID"] = courseNo;
                        DRh["CCName"] = _ccName;
                        DRh["CourseName"] = _courseName[i - 1];
                        DRh["Hole"] = h + 1;
                        DRh["WebPar"] = par[h];
                        _DTtempResult.Rows.Add(DRh.ItemArray);


                        // 티마다 생성하는 티거리
                        for (int t = 0; t < 5; t++)
                        {
                            DataRow DRt = _DTtempResult.NewRow();
                            DRt["CCID"] = courseNo;
                            DRt["CCName"] = _ccName;
                            DRt["CourseName"] = _courseName[i - 1];
                            DRt["Hole"] = h + 1;
                            DRt["Tee"] = tee[t];
                            DRt["WebDistance"] = tDis[t, h];
                            _DTtempResult.Rows.Add(DRt.ItemArray);
                        }
                    }
                }
            }
            // 데이터테이블이 다 만들어지면 시험용 익스포트
            // Class.Excel EX = new Class.Excel();
            // EX.Excel2007Export(_DTresult, _resultPath + "WebData.xlsx");            
        }


        #region ---- 이전 소스 코드 ----

        // 스타트 버튼 클릭 이벤트
        private void BTNstart_Click(object sender, EventArgs e)
        {
            // 패스워드 변수 입력
            _password = TXTpassword.Text;
            LOG.Write("GS 패스워드 : " + _password);

            // GS 타입 반영 << 향후 수정 필요
            GS_TYPE = CBBgsType.Text;

            // GS 타입에 따른 어플리케이션 네임 배정
            switch (GS_TYPE)
            {
                case "VISION":
                default:
                    DEFAULT_APP_NAME = "GolfZonCilent";
                    _imagePath = Application.StartupPath + @"\Images\Vision\";
                    _left = 0;
                    break;


                case "TWOVISION":
                    // 투비전은 터치모니터 프로세스를 띄어야 정상적인 입력이 가능
                    DEFAULT_APP_NAME = "GolfzonTouchMonitor";
                    _imagePath = Application.StartupPath + @"\Images\Twovision\";

                    // 풀스크린 캡쳐 시 주스크린만 찍음
                    _resX = 1920;
                    _resY = 1080;
                    _left = -1920;
                    break;
            }

            // 작업홀을 읽어옴
            string[] tempWorkHoles = TXThole.Text.Split(',');
            Array.Resize(ref _workHoles, tempWorkHoles.Length);

            for (int i = 0; i < tempWorkHoles.Length; i++)
            {
                _workHoles[i] = tempWorkHoles[i].Trim();
                LOG.Write("_workHole = " + _workHoles[i].ToString());
            }

            // 반복횟수 읽어옴
            int repeatCount = int.Parse(TXTrepeat.Text);
            LOG.Write("repeatCount = " + repeatCount);


            // 매크로 시작 (타이틀 화면에서)            
            switch (GS_TYPE)
            {
                case "VISION":
                default:
                    LOG.Write("InitVision() 실행");
                    InitVision();
                    break;


                case "TWOVISION":
                    LOG.Write("InitTwovision() 실행");
                    InitTwovision();
                    break;
            }


            // 반복횟수 만큼 반복
            for (int i = 0; i < repeatCount; i++)
            {
                // 작업홀 만큼 반복
                for (int j = 0; j < _workHoles.Length; j++)
                {
                    LOG.Write("AdCaptureSelectedHoles 실행, repeatCount = " + i.ToString()
                        + " ,workHoles = " + _workHoles[j].ToString());


                    // _workHoles[j]가 숫자인지 문자열(17-18)인지 판단하여 함수를 구분
                    // 숫자로 변환이 가능할 경우 해당홀+다음홀에 나오는 로직 태움
                    if (int.TryParse(_workHoles[j], out int workHole))
                    {
                        int[] workHoles = new int[1];
                        workHoles[0] = workHole;

                        // GS 타입에 따라 다른 메소드 실행
                        switch (GS_TYPE)
                        {
                            case "VISION":
                            default:
                                AdCaptureSelectedHolesVision(workHoles);
                                break;


                            case "TWOVISION":
                                AdCaptureSelectedHolesTwovision(workHoles);
                                break;
                        }

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
                            switch (GS_TYPE)
                            {
                                case "VISION":
                                default:
                                    AdCaptureSelectedHolesVision(workHoles);
                                    break;

                                case "TWOVISION":
                                    AdCaptureSelectedHolesTwovision(workHoles);
                                    break;
                            }

                        }
                        catch (Exception error)
                        {
                            LOG.Write(error.Message);
                            Application.ExitThread();
                            Environment.Exit(0);
                        }
                    }

                    LOG.Write("AdCaptureSelectedHoles 완료, repeatCount = " + i.ToString()
                        + " ,workHoles = " + _workHoles[j].ToString());
                }

                _repeatCount++;
            }


            Thread.Sleep(3000);
            LOG.Write("코드 실행 완료");


            AutoItX.WinActivate(Application.ProductName);
            MessageBox.Show("실행 완료!!!");
        }


        // 게스트 로그인 후 모드선택까지 이동
        private void InitVision()
        {
            try
            {
                // 창 활성화
                LOG.Write("어플리케이션 창 활성화 시도");
                AutoItX.WinActivate(DEFAULT_APP_NAME);


                // 어플리케이션 핸들을 받아오지 못하면 강제 종료
                string handle = AutoItX.WinGetHandleAsText(DEFAULT_APP_NAME);
                LOG.Write("handle = " + handle);

                if (handle == "0x00000000")
                {
                    LOG.Write("어플리케이션 창 핸들을 받아오지 못함");
                    throw new Exception("어플리케이션 창 핸들을 받아오지 못함");
                }
                LOG.Write("어플리케이션 창 활성화 성공");


                // 창 활성화까지 대기
                Thread.Sleep(3000);


                // 아무키나 입력해서 비번 창을 뛰움
                AutoItX.Send("{ENTER}");


                // 비번 사용 시 
                if (CHKusePassword.Checked)
                {
                    // 비번 창 활성화까지 대기
                    Thread.Sleep(2000);


                    // 비번 창 텍스트 박스 클릭 (혹시 포커스가 안가있을지도 모를 사태 대비)
                    AutoItX.MouseClick("LEFT", 881, 436);
                    Thread.Sleep(3000);


                    // 패스워드 입력
                    AutoItX.Send(_password);
                    Thread.Sleep(1000);
                    AutoItX.Send("{Enter}");
                    LOG.Write("패스워드 입력 완료");
                }


                // 플레이어 설정 전환까지 대기
                AutoItX.MouseMove(_resX, _resY);
                Thread.Sleep(3000);


                // 정상적으로 플레이어 설정 진입했는지 확인절차                
                Bitmap playerConfig = CapturePartialScreen(317, 289, 169, 38, _tempPath + "플레이어설정확인");
                LOG.Write("플레이어 설정 진입 확인");

                if (ImageToString(playerConfig, "kor").Trim() != "플레이어 설정")
                {
                    LOG.Write("플레이어 설정이 아님. 매크로 종료");
                    throw new Exception("플레이어 설정 진입 실패");
                }


                // 게스트 등록 버튼 클릭
                AutoItX.MouseClick("LEFT", 890, 225);
                Thread.Sleep(2000);


                // 다음 버튼 클릭
                AutoItX.MouseClick("LEFT", 1477, 924);


                // 모드선택 - 대회정보를 수신할때까지 넉넉히 대기 
                LOG.Write("모드선택 진입");
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


        // 투비전 게스트 로그인 후 모드선택까지 이동
        private void InitTwovision()
        {
            try
            {
                // 혹시 모를 화면가림 대비를 위해 golfzonclient에 포커스를 한번 줌
                AutoItX.WinActivate("GolfZonClient");
                Thread.Sleep(1000);


                // 창 활성화            
                AutoItX.WinActivate(DEFAULT_APP_NAME);


                // 어플리케이션 핸들을 받아오지 못하면 강제 종료
                string handle = AutoItX.WinGetHandleAsText(DEFAULT_APP_NAME);
                LOG.Write("handle = " + handle);

                if (handle == "0x00000000")
                {
                    LOG.Write("어플리케이션 창 핸들을 받아오지 못함");
                    throw new Exception("어플리케이션 창 핸들을 받아오지 못함");
                }
                LOG.Write("어플리케이션 창 활성화 성공");


                // 창 활성화까지 대기
                Thread.Sleep(3000);
                LOG.Write("타이틀 초기상태 확인");


                // 비밀번호 사용이 체크되어 있으면
                if (CHKusePassword.Checked)
                {
                    // 알트탭을 하면 이미 비밀번호 창이 떠 있고 포커스가 있는 상태                
                    // 초기상태 확인

                    CheckVKPasswordAndInput();
                    LOG.Write("패스워드 입력 완료");


                    // 홀/시간설정 팝업 확인 누름
                    AutoItX.MouseClick("LEFT", 840, 690);
                }
                // 비밀번호 사용이 체크되어 있지 않으면 시작 버튼만 누름
                else
                {
                    AutoItX.MouseClick("LEFT", 960, 670);
                }


                // 플레이어 설정 전환까지 대기                
                Thread.Sleep(3000);


                // 정상적으로 플레이어 설정 진입했는지 확인절차                
                LOG.Write("플레이어 설정 진입 확인");

                Bitmap playerConfig = CapturePartialScreen(264, 134, 264, 68, _tempPath + "플레이어설정확인");
                Bitmap playerConfig2 = new Bitmap(_imagePath + "PlayerConfig.png");

                if (CompareImages(playerConfig, playerConfig2) < 98)
                {
                    throw new Exception("플레이어 설정 진입 실패");
                }


                // 게스트 등록 버튼 클릭
                AutoItX.MouseClick("LEFT", 894, 273);
                Thread.Sleep(3000);


                // 다음 버튼 클릭
                AutoItX.MouseClick("LEFT", 1807, 436);


                // 모드선택 - 대회정보를 수신할때까지 넉넉히 대기 
                LOG.Write("모드선택 진입");
                Thread.Sleep(5000);
            }
            catch (Exception error)
            {
                LOG.Write(error.Message);
                AutoItX.WinActivate(Application.ProductName);
                MessageBox.Show(error.Message);

                Application.ExitThread();
                Environment.Exit(0);
            }
        }


        // 인게임 내 홀스킵 이동 광고 캡쳐
        private void AdCaptureSelectedHolesVision(int[] holes)
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
                LOG.Write("CC선택 진입");
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
                LOG.Write("라운드설정 진입");
                Thread.Sleep(3000);


                // 기타설정 클릭
                AutoItX.MouseClick("LEFT", 1594, 487);
                Thread.Sleep(2000);


                // 시작홀 선택                
                ClickStartHole(firstHole);


                // 라운드시작 클릭 
                LOG.Write("라운드 시작 클릭");
                AutoItX.MouseClick("LEFT", 1499, 914);
                Thread.Sleep(2000);


                // 서비스 이용 상세내역 팝업창
                LOG.Write("서비스 이용 상세내역 팝업창 진입");
                AutoItX.MouseClick("LEFT", 833, 580);
                Thread.Sleep(1000);


                // 로딩 광고 레코딩 시작
                string videoFileName = firstHole.ToString() + "H 로딩광고_" + _repeatCount;
                RecordToNextHole(firstHole, videoFileName);
                Thread.Sleep(3000);


                foreach (int hole in holes)
                {
                    LOG.Write(hole.ToString() + "H 진입완료");
                    AutoItX.WinActivate(DEFAULT_APP_NAME);

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
                        LOG.Write("홀스킵 팝업창 진입");
                        AutoItX.Send("{ENTER}");
                        Thread.Sleep(2000);


                        // 스코어보드 광고 촬영
                        CaptureFullScreen(_resultPath + hole.ToString() + "H 스코어보드 광고_" + _repeatCount);


                        // 비디오 레코딩 시작
                        videoFileName = (hole + 1).ToString() + "H 로딩광고_" + _repeatCount;
                        RecordToNextHole(hole + 1, videoFileName);
                        Thread.Sleep(3000);
                    }
                }

                // 종료로직 수행
                IngameQuitVision();


                // 모드선택으로 돌아올때까지 대기
                Thread.Sleep(15000);


                // 비번 사용 시 비밀번호 입력
                if (CHKusePassword.Checked)
                {
                    // 홀/시간 설정 진입
                    LOG.Write("모드선택 홀/시간설정 팝업 진입확인");

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
            }

            catch (Exception error)
            {
                LOG.Write(error.Message);
                AutoItX.WinActivate(Application.ProductName);
                MessageBox.Show(error.Message);

                Application.ExitThread();
                Environment.Exit(0);
            }
        }


        // 인게임 내 홀스킵 이동 광고 캡쳐
        private void AdCaptureSelectedHolesTwovision(int[] holes)
        {
            try
            {
                // 홀은 시작홀
                int firstHole = holes[0];


                // 모드선택 광고팝업 유무 확인
                LOG.Write("모드선택 광고팝업 유무 확인");

                Bitmap title = CapturePartialScreen(423, 137, 223, 37, _tempPath + "골프존추천프로모션");
                Bitmap title2 = new Bitmap(_imagePath + "ModeSelectPromotionPop.png");


                // 광고창이 떠 있으면 닫기를 눌러줌
                if (CompareImages(title, title2) >= 98)
                {
                    AutoItX.MouseClick("LEFT", 957, 873);
                    Thread.Sleep(3000);
                }


                // 스트로크 선택
                AutoItX.MouseClick("LEFT", 704, 754);
                Thread.Sleep(3000);


                // 스트로크 상태에서 다음
                AutoItX.MouseClick("LEFT", 1807, 432);


                // CC 선택 - 넉넉히 대기
                LOG.Write("CC선택 진입");
                Thread.Sleep(5000);


                // CC 검색 클릭
                AutoItX.MouseClick("LEFT", 1534, 151);
                Thread.Sleep(3000);


                // 투비전은 CC명에서 공백 제거
                string ccName = Regex.Replace(TXTccName.Text, @"\s", "");


                // CC 입력
                InputVirtualKeyBoard(ccName);


                // CC 검색시간 대기
                Thread.Sleep(3000);


                // CC 선택에서 다음
                AutoItX.MouseClick("LEFT", 1807, 432);
                Thread.Sleep(3000);


                // 코스매니저와 볼설정 팝업창
                AutoItX.MouseClick("LEFT", 834, 743);


                // 라운드 설정으로 이동
                LOG.Write("라운드설정 진입");


                // 팝업창으로 인한 여유
                Thread.Sleep(4000);


                // 기타설정 클릭
                AutoItX.MouseClick("LEFT", 1561, 871);
                Thread.Sleep(2000);


                // 시작홀 선택                
                ClickStartHole(firstHole);


                // 라운드시작 클릭 
                LOG.Write("라운드 시작 클릭");
                AutoItX.MouseClick("LEFT", 1807, 432);
                Thread.Sleep(2000);


                // 광고 레코딩 시작
                string videoFileName = firstHole.ToString() + "H 로딩광고_" + _repeatCount;
                RecordToNextHole(firstHole, videoFileName);


                // 홀 진입 상태                
                Thread.Sleep(3000);

                foreach (int hole in holes)
                {
                    LOG.Write(hole.ToString() + "H 진입완료");


                    // 터치스크린 리프레시 (안하면 터치가 헛돌음)
                    AutoItX.WinActivate(DEFAULT_APP_NAME);


                    // 나스모 광고 촬영
                    if (CHKnasmo.Checked)
                    {
                        // 메뉴 옆으로 이동
                        AutoItX.MouseClick("LEFT", 1820, 800);
                        Thread.Sleep(2000);


                        // 나스모 클릭
                        AutoItX.MouseClick("LEFT", 810, 680);
                        Thread.Sleep(3000);
                        CaptureFullScreen(_resultPath + hole.ToString() + "H 나스모 광고_" + _repeatCount);


                        // 닫기 버튼 클릭
                        AutoItX.MouseClick("LEFT", 950, 850);
                        Thread.Sleep(3000);
                    }

                    // 18홀일 경우 이후의 액션을 스킵하고 바로 종료
                    if (hole != 18)
                    {
                        // 홀스킵 사용
                        AutoItX.MouseClick("LEFT", 1820, 800);
                        Thread.Sleep(2000);
                        AutoItX.MouseClick("LEFT", 1610, 680);
                        Thread.Sleep(2000);


                        // 홀스킵 팝업창 진입
                        LOG.Write("홀스킵 팝업창 진입");
                        AutoItX.MouseClick("LEFT", 840, 690);
                        Thread.Sleep(2000);


                        // 스코어보드 광고 촬영
                        CaptureFullScreen(_resultPath + hole.ToString() + "H 스코어보드 광고_" + _repeatCount);


                        // 홀간 광고 레코딩 시작
                        videoFileName = (hole + 1).ToString() + "H 로딩광고_" + _repeatCount;
                        RecordToNextHole(hole + 1, videoFileName);
                        Thread.Sleep(3000);
                    }
                }

                // 종료로직 수행
                IngameQuitTwovision();


                // 모드선택으로 돌아올때까지 대기
                Thread.Sleep(15000);


                // 비밀번호 사용 시 비번입력
                if (CHKusePassword.Checked)
                {
                    // 홀/시간 설정 진입
                    LOG.Write("모드선택 홀/시간설정 팝업 진입확인");

                    // 비번 화상키보드 확인 및 입력
                    CheckVKPasswordAndInput();

                    // 홀/시간설정 팝업 닫음
                    AutoItX.MouseClick("LEFT", 837, 689);
                }

                LOG.Write("모드선택 진입");

                // 대회 정보 수신시간 확보
                Thread.Sleep(5000);
            }

            catch (Exception error)
            {
                LOG.Write(error.Message);
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
            ImageUpScale(modeSelect);
            LOG.Write("모드선택 진입 확인");


            // 모드선택 -> 모드선랙으로 읽음 *TS 엔진 한계
            if (ImageToString(modeSelect, "kor").Trim() != "모드 선태")
            {
                LOG.Write("모드선택 진입 실패. 매크로 종료");
                throw new Exception("모드선택 진입 실패");
            }
        }


        // 인게임 종료 팝업창 클릭 버그 해결
        private void IngameQuitVision()
        {
            try
            {
                LOG.Write("게임종료 로직 수행");


                // 클릭 삑사리 대비용
                AutoItX.WinActivate(DEFAULT_APP_NAME);


                // 게임종료
                AutoItX.Send("{ESC}");
                Thread.Sleep(3000);


                // 비번 사용 시 비밀번호 입력
                if (CHKusePassword.Checked)
                {
                    // 관리자 비번 진입
                    LOG.Write("게임종료 관리자비번 진입");
                    AutoItX.Send(_password);
                    Thread.Sleep(3000);
                    AutoItX.Send("{ENTER}");
                    Thread.Sleep(3000);
                }


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
                    LOG.Write("유지종료 클릭 = " + timeOut.ToString());

                    Thread.Sleep(1000);
                }
            }
            catch (Exception error)
            {
                LOG.Write(error.Message);
                AutoItX.WinActivate(Application.ProductName);
                MessageBox.Show(error.Message);

                Application.ExitThread();
                Environment.Exit(0);
            }
        }


        // 
        private void IngameQuitTwovision()
        {
            try
            {
                LOG.Write("게임종료 로직 수행");


                // 이게 없으면 마우스 동작 및 화상키보드가 먹지 않음
                AutoItX.WinActivate(DEFAULT_APP_NAME);


                // 게임종료
                AutoItX.MouseClick("LEFT", 1850, 80);
                Thread.Sleep(3000);


                // 비밀번호 사용 시 비번 입력
                if (CHKusePassword.Checked)
                {
                    // 관리자 비번 진입
                    LOG.Write("게임종료 관리자비번 진입");


                    // 비번 화상키보드 확인 & 입력
                    CheckVKPasswordAndInput();


                    // 비번 팝업창 확인 버튼 클릭
                    AutoItX.MouseClick("LEFT", 840, 690);
                    Thread.Sleep(3000);
                }


                // 유지종료 (한번하면 잘 안됨)
                int tolerance = 100;
                int timeOut = 0;

                // 종료팝업이 없어질때까지 무한 반복
                while (tolerance >= 98)
                {
                    Bitmap screenShot = CapturePartialScreen(808, 365, 295, 135, _tempPath + "게임종료팝업");
                    Bitmap img = new Bitmap(_imagePath + "IngameQuitPop.png");
                    tolerance = CompareImages(screenShot, img);
                    timeOut++;

                    AutoItX.MouseClick("LEFT", 960, 640);
                    LOG.Write("유지종료 클릭 = " + timeOut.ToString());

                    Thread.Sleep(1000);
                }
            }
            catch (Exception error)
            {
                LOG.Write(error.Message);
                AutoItX.WinActivate(Application.ProductName);
                MessageBox.Show(error.Message);

                Application.ExitThread();
                Environment.Exit(0);
            }
        }


        // 시작홀을 받아서 라운드설정-홀선택에서 마우스 좌표를 찍게 해줌
        private void ClickStartHole(int startHole)
        {
            LOG.Write("시작홀 = " + startHole.ToString());


            // 칼럼번호 1홀 = 1
            int column = startHole;
            int clickX = 0;
            int clickY = 0;


            // 1홀 스타트 클릭 위치
            switch (GS_TYPE)
            {
                case "REAL":
                    break;

                case "VISION":
                default:

                    clickX = 950;
                    clickY = 600;
                    break;

                case "TWOVISION":
                    clickX = 1010;
                    clickY = 750;
                    break;
            }


            // 스타트 홀이 아랫줄일 경우 Y좌표 이동
            if (startHole > 9)
            {
                switch (GS_TYPE)
                {
                    case "REAL":
                        break;

                    default:
                    case "VISION":
                        clickY = 685;
                        break;

                    case "TWOVISION":
                        clickY = 815;
                        break;
                }

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
            // 좌표 선언 (비전 & 투비전 공용)
            int x = 75;
            int y = 18;
            int width = 47;
            int height = 53;

            switch (GS_TYPE)
            {
                case "REAL":
                    break;

                case "VISION":
                default:
                    break;

                case "TWOVISION":
                    // 주스크린을 찍으려면 x 좌표 수정 필요
                    x -= 1920;
                    break;
            }


            Bitmap img = CapturePartialScreen(x, y, width, height, _tempPath + "홀번호");
            List<string> fileNames = new List<string>();
            int holeNo = 0;


            // 비전의 경우 VisionHoleImg 폴더 사용 ★★★ 다른 GS에서는 다른 폴더를 이용해야 함
            DirectoryInfo di = new DirectoryInfo(_imagePath);

            foreach (FileInfo file in di.GetFiles())
            {
                // 홀 번호 대상 이미지로만 검색
                if (!file.Name.Contains("HoleNo_"))
                {
                    continue;
                }

                Bitmap holeImg = new Bitmap(file.FullName);

                // 이미지 정합성이 98% 이상일경우 해당 홀로 판정
                LOG.Write("홀번호.png" + " vs " + file.Name);
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


            // 홀 이미지가 매칭될때까지 0.5초 단위로 확인
            while (hole != currentHole)
            {
                LOG.Write("홀 이미지 매칭 중 : currentHole = " + currentHole.ToString());

                Thread.Sleep(500);
                overTime += 1;
                currentHole = GetHoleNumber();


                // 만약 120초간 대기에도 매칭이 되지 않으면 강제 종료
                if (overTime > 120)
                {
                    LOG.Write("매크로 종료 : 홀 이미지 매칭 대기 초과");
                    throw new Exception("홀 이미지 매칭 대기 초과");
                }
            }

            return;
        }


        // 범용 이미지 체크 로직 (체크에 실패할 경우 어플리케이션 강제 종료
        private void CommonImageCheck(int x, int y, int width, int height, string filePath, int tolerance)
        {
            Bitmap screenShot = CapturePartialScreen(x, y, width, height, _tempPath + "imgChk1");
            Bitmap img = new Bitmap(filePath);

            LOG.Write("imgChk1 vs " + filePath + " , tolerance = " + tolerance.ToString());

            // 이미지가 톨레랑스보다 낮을 경우 어플 종료
            if (CompareImages(screenShot, img) < tolerance)
            {
                LOG.Write("이미지 체크 실패");
                Application.ExitThread();
                Environment.Exit(0);
            }
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


        // 녹화 템프 삭제
        private void DeleteRecordingTemp()
        {
            LOG.Write("레코딩 템프 삭제");

            foreach (string file in _inputImageSequence)
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception error)
                {
                    LOG.Write(error.Message);
                }

            }
        }


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


        #endregion


        #region ---- 캡쳐 메소드들 ----

        // 비디오 레코드
        private void RecordVideo(int term, string fileName)
        {
            LOG.Write("비디오 레코드 시작");

            DateTime targetTime = DateTime.Now.AddMilliseconds(term);

            // 초당 8 ~ 20프레임이 한계
            while (DateTime.Now <= targetTime)
            {
                Recording();
            }


            // 생성된 이미지 파일 수로 영상의 fps를 계산
            double fps = (double)_fileCount / ((double)term / 1000);
            fps = Math.Round(fps);

            LOG.Write("_fileCount = " + _fileCount.ToString("N0")
                + " ,term = " + term.ToString("N0")
                + "ms ,fps = " + fps.ToString());

            if (fps != 0)
            {
                EncodingVideo((int)fps, fileName);
            }
            else
            {
                LOG.Write("계산된 fps가 0이어서 영상을 인코딩할 수 없음.");
            }

        }



        // 비디오 녹화 <- 동기화 처리
        private void Recording()
        {
            using (Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.CopyFromScreen(_left, _top, 0, 0, bitmap.Size);
                }

                string fileName = _tempPath + "videoTemp_" + _fileCount + ".png";
                bitmap.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                _inputImageSequence.Add(fileName);
                _fileCount++;


                bitmap.Dispose();
            }
        }



        // 다음홀까지 비동기 레코딩 실행
        private void RecordToNextHole(int hole, string fileName)
        {
            _recordingStartTime = DateTime.Now;


            // 화면녹화를 비동기로 진행
            LOG.Write("ThreadRecording 실행 : hole = " + hole.ToString() + ", fileName = " + fileName);
            _tRecording = new Thread(ThreadRecording);
            _tRecording.Start();


            int currentHole = GetHoleNumber();
            int chkCount = 0;


            // 홀 이미지가 매칭될때까지 0.5초 단위로 확인
            while (hole != currentHole)
            {
                LOG.Write("홀 이미지 매칭 중 : currentHole = " + currentHole.ToString());

                Thread.Sleep(500);
                chkCount += 1;
                currentHole = GetHoleNumber();


                // 만약 120회 수행에도 매칭이 되지 않으면 강제 종료
                if (chkCount > 120)
                {
                    LOG.Write("매크로 종료 : 홀 이미지 매칭 대기 초과");

                    _tRecording.Abort();
                    _tRecording.Join();

                    throw new Exception("홀 이미지 매칭 대기 초과");
                }
            }

            // 화면녹화 종료
            _tRecording.Abort();
            _tRecording.Join();
            LOG.Write("ThreadRecording 종료");


            // 비동기 인코딩 시작 --> 끝남을 기다리지 않음
            AsyncEncodingVideo(fileName);

            return;
        }



        // 비디오 녹화 <- 비동기 스레딩 방식
        private void ThreadRecording()
        {
            while (true)
            {
                using (Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.CopyFromScreen(_left, _top, 0, 0, bitmap.Size);
                    }

                    string fileName = _tempPath + "videoTemp_" + _fileCount + ".png";
                    bitmap.Save(fileName, System.Drawing.Imaging.ImageFormat.Png);
                    _inputImageSequence.Add(fileName);
                    _fileCount++;


                    bitmap.Dispose();
                }
            }
        }



        // 비디오 파일 저장
        private void EncodingVideo(int fps, string fileName)
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

                    LOG.Write("비디오 인코딩 시작");
                    LOG.Write("파일 카운트 : " + _fileCount.ToString("N0"));

                    // 이미지 시퀀스 비디오 인코딩
                    foreach (string imageLoc in _inputImageSequence)
                    {
                        Bitmap imageFrame = Image.FromFile(imageLoc) as Bitmap;
                        vfw.WriteVideoFrame(imageFrame);
                        imageFrame.Dispose();
                    }

                    LOG.Write("비디오 인코딩 완료");


                    // 레코딩 템프 삭제
                    DeleteRecordingTemp();


                    // 파일카운트, 저장변수 초기화
                    _fileCount = 1;
                    _inputImageSequence.Clear();


                    vfw.Close();
                }

                LOG.Write("비디오파일 저장 완료 : " + _resultPath + fileName + ".mp4");
            }
            catch (Exception e)
            {
                LOG.Write(e.Message);
            }
        }



        // 비동기 스레딩 방식의 비디오 저장 전처리
        private void AsyncEncodingVideo(string fileName)
        {
            // 현재까지 러닝타임 계산
            TimeSpan recTime = DateTime.Now - _recordingStartTime;
            double term = recTime.TotalMilliseconds;


            // 생성된 이미지 파일 수로 영상의 fps를 계산
            double fps = (double)_fileCount / (term / 1000);
            fps = Math.Round(fps);


            LOG.Write("_fileCount = " + _fileCount.ToString("N0")
                + " ,term = " + term.ToString("N0")
                + "ms ,fps = " + fps.ToString());


            // fps가 0이 아닐때만 인코딩 작업
            if (fps != 0)
            {
                EncodingVideo((int)fps, fileName);
            }
            else
            {
                LOG.Write("계산된 fps가 0이어서 영상을 인코딩할 수 없음.");
            }
        }



        // 전체 스크린 캡쳐
        private void CaptureFullScreen(string fileName)
        {
            LOG.Write("스크린 샷 캡쳐 시도 (" + fileName + ") _resX : " + _resX.ToString());

            try
            {
                Bitmap screenshot = new Bitmap(_resX, _resY, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                Graphics gr = Graphics.FromImage(screenshot);
                gr.CopyFromScreen(_left, _top, 0, 0, screenshot.Size);
                screenshot.Save(fileName + ".png", System.Drawing.Imaging.ImageFormat.Png);
                LOG.Write("스크린 샷 캡쳐 성공 (" + fileName + ")");
            }
            catch (Exception error)
            {
                LOG.Write("스크린 샷 캡쳐 실패 = " + error.Message);
            }
        }



        // 부분 스크린 캡쳐
        private Bitmap CapturePartialScreen(int x, int y, int width, int height, string fileName)
        {
            LOG.Write("스크린 샷 캡쳐 시도 (" + fileName + ")");

            try
            {
                Bitmap screenshot = new Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                Graphics gr = Graphics.FromImage(screenshot);
                gr.CopyFromScreen(x, y, 0, 0, screenshot.Size);
                screenshot.Save(fileName + ".png", System.Drawing.Imaging.ImageFormat.Png);
                LOG.Write("스크린 샷 캡쳐 성공 (" + fileName + ")");
                return screenshot;
            }
            catch (Exception error)
            {
                LOG.Write("스크린 샷 캡쳐 실패 = " + error.Message);
            }

            return null;
        }
        #endregion


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
                LOG.Write("이미지 그레이스케일 변환 완료");

                // 이진화
                Cv2.Threshold(gray, bin, 0, 255, ThresholdTypes.Otsu);


                // 이진화 파일 저장
                Cv2.ImWrite(_tempPath + "grayscale.png", gray);
                Cv2.ImWrite(_tempPath + "bin.png", bin);

                LOG.Write("이미지 이진화 변환 완료");
                return BitmapConverter.ToBitmap(bin);
            }
            catch (Exception error)
            {
                LOG.Write(error.Message);
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
            LOG.Write("이미지 업스케일링 변환 완료");

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
            LOG.Write("이미지 다운스케일링 변환 완료");

            return downScale;
        }



        // 이미지 역상(색상 반전)
        private Bitmap ImageReverse(Bitmap img)
        {
            Mat src = BitmapConverter.ToMat(img);
            Mat dst = new Mat();

            Cv2.BitwiseNot(src, dst);
            Cv2.ImWrite(_tempPath + "reverse.png", dst);

            LOG.Write("이미지 역상 변환 완료");

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
                LOG.Write("이미지 비트맵 비교 결과 : " + result.ToString("N2") + "% 일치");

                return (int)result;
            }

            LOG.Write("이미지 비트맵 비교 결과 : 이미지 사이즈가 맞지 않음");

            return -1;
        }



        // OCR로 이미지 파일을 받아 문자로 변환
        private string ImageToString(Bitmap img, string language)
        {
            try
            {
                using (var engine = new TesseractEngine(@"./tessdata", language, EngineMode.TesseractOnly))
                {
                    using (var page = engine.Process(img))
                    {
                        LOG.Write("이미지 Tesseract 결과 : " + page.GetText());
                        return page.GetText();
                    }
                }
            }
            catch (Exception error)
            {
                LOG.Write(error.Message);
            }

            return null;
        }
        #endregion


        #region ---- 화상키보드 관련 메소드 ----

        // 한글키를 눌러주는 키보드 이벤트
        private void PressHangulKey()
        {
            keybd_event((byte)Keys.HanguelMode, 0, 0x00, 0);
            keybd_event((byte)Keys.HanguelMode, 0, 0x02, 0);
        }


        // 투비전용 비번 화상키보드 검사
        private void CheckVKPasswordAndInput()
        {
            LOG.Write("화상키보드 비밀번호 입력 진입");

            // 비번 화상키보드 확인
            Bitmap vkPass = CapturePartialScreen(375, 494, 111, 74, _tempPath + "화상키보드비번");
            Bitmap vkPass2 = new Bitmap(_imagePath + "TitlePasswordPop.png");
            if (CompareImages(vkPass, vkPass2) < 98)
            {
                throw new Exception("화상키보드 비번 진입실패");
            }

            // 화상키보드 비번 입력
            Thread.Sleep(1000);
            InputVirtualKeyBoard(_password);
            Thread.Sleep(3000);
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

                LOG.Write("키보드 입력 신호 전달 = " + sendText);
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
        private void InputVirtualKeyBoard(string text)
        {
            // 디스플레이 포커스 유지
            AutoItX.WinActivate(DEFAULT_APP_NAME);


            LOG.Write("문자열 = " + text);

            foreach (char ch in text)
            {
                LOG.Write("문자 = " + ch.ToString());

                // 한글이 아니면 분해를 하지 않음
                if (CheckString(ch) != "K")
                {
                    LOG.Write("자소 분해를 하지 않음");

                    // 영어일 경우 한영키를 전환
                    if (CheckString(ch) == "E")
                    {
                        // 영키 전환
                        AutoItX.MouseClick("LEFT", 1400, 920);
                        Thread.Sleep(200);

                        TouchVirtualKeyBoard(ch);

                        // 한키 전환
                        AutoItX.MouseClick("LEFT", 1400, 920);
                        Thread.Sleep(200);
                    }
                    // 숫자 또는 공백일 경우 전환없이 그냥 키보드 입력
                    else
                    {
                        TouchVirtualKeyBoard(ch);
                    }
                }
                // 한글일 경우 자소 분해 
                else
                {
                    int temp = Convert.ToUInt16(ch);
                    int nUniCode = temp - 0xAC00;

                    // 초성 중성 종성 변수
                    int cho = nUniCode / (21 * 28);
                    nUniCode = nUniCode % (21 * 28);

                    int jung = nUniCode / 28;
                    int jong = nUniCode % 28;

                    char[] jaso = new char[3];
                    jaso[0] = CHOSUNG_TABLE[cho];
                    jaso[1] = JUNGSUNG_TABLE[jung];
                    jaso[2] = JONGSUNG_TABLE[jong];

                    for (int i = 0; i < jaso.Length; i++)
                    {
                        TouchVirtualKeyBoard(jaso[i]);
                        LOG.Write("자소 = " + jaso[i].ToString());
                    }
                }
            }

            // 마지막 입력에 ENTER를 클릭해줌 (send{ENTER} 안먹힘)
            Thread.Sleep(1000);
            LOG.Write("마지막 [ENTER]키 클릭");
            AutoItX.MouseClick("LEFT", 1460, 820);
        }


        // 자소를 받아 터치스크린을 클릭
        private void TouchVirtualKeyBoard(char ch)
        {
            LOG.Write("입력받은 문자 ch = " + ch.ToString());

            // 클릭 x, y 좌표 초기화
            int x = 0;
            int y = 0;

            // 시프트키가 필요한지
            bool isNeedShift = false;

            // 화상키보드 좌표와 연동해 클릭
            switch (ch)
            {
                case 'ㄱ':
                case 'r':
                    x = 774;
                    y = 692;
                    break;

                case 'ㄲ':
                    isNeedShift = true;
                    x = 774;
                    y = 692;
                    break;

                case 'ㄴ':
                case 's':
                    x = 645;
                    y = 770;
                    break;

                case 'ㄷ':
                case 'e':
                    x = 687;
                    y = 704;
                    break;

                case 'ㄸ':
                    isNeedShift = true;
                    x = 687;
                    y = 704;
                    break;

                case 'ㄹ':
                case 'f':
                    x = 820;
                    y = 773;
                    break;

                case 'ㅁ':
                case 'a':
                    x = 591;
                    y = 779;
                    break;

                case 'ㅂ':
                case 'q':
                    x = 545;
                    y = 689;
                    break;

                case 'ㅃ':
                    isNeedShift = true;
                    x = 545;
                    y = 689;
                    break;

                case 'ㅅ':
                case 't':
                    x = 861;
                    y = 692;
                    break;

                case 'ㅆ':
                    isNeedShift = true;
                    x = 861;
                    y = 692;
                    break;

                case 'ㅇ':
                case 'd':
                    x = 727;
                    y = 767;
                    break;

                case 'ㅈ':
                case 'w':
                    x = 609;
                    y = 686;
                    break;

                case 'ㅉ':
                    isNeedShift = true;
                    x = 609;
                    y = 686;
                    break;

                case 'ㅊ':
                case 'c':
                    x = 750;
                    y = 846;
                    break;

                case 'ㅋ':
                case 'z':
                    x = 580;
                    y = 852;
                    break;

                case 'ㅌ':
                case 'x':
                    x = 664;
                    y = 855;
                    break;

                case 'ㅍ':
                case 'v':
                    x = 825;
                    y = 851;
                    break;

                case 'ㅎ':
                case 'g':
                    x = 889;
                    y = 767;
                    break;

                case 'ㅏ':
                case 'k':
                    x = 1126;
                    y = 779;
                    break;

                case 'ㅐ':
                case 'o':
                    x = 1161;
                    y = 687;
                    break;

                case 'ㅑ':
                case 'i':
                    x = 1096;
                    y = 698;
                    break;

                case 'ㅒ':
                    isNeedShift = true;
                    x = 1161;
                    y = 687;
                    break;

                case 'ㅓ':
                case 'j':
                    x = 1053;
                    y = 762;
                    break;

                case 'ㅔ':
                case 'p':
                    x = 1260;
                    y = 681;
                    break;

                case 'ㅕ':
                case 'u':
                    x = 1015;
                    y = 690;
                    break;

                case 'ㅖ':
                    isNeedShift = true;
                    x = 1260;
                    y = 681;
                    break;

                case 'ㅗ':
                case 'h':
                    x = 982;
                    y = 777;
                    break;

                case 'ㅘ':
                    // 겹모음은 메소드 재귀호출
                    TouchVirtualKeyBoard('ㅗ');
                    TouchVirtualKeyBoard('ㅏ');
                    return;

                case 'ㅙ':
                    TouchVirtualKeyBoard('ㅗ');
                    TouchVirtualKeyBoard('ㅐ');
                    return;

                case 'ㅚ':
                    TouchVirtualKeyBoard('ㅗ');
                    TouchVirtualKeyBoard('ㅣ');
                    return;

                case 'ㅛ':
                case 'y':
                    x = 931;
                    y = 686;
                    break;

                case 'ㅜ':
                case 'n':
                    x = 978;
                    y = 859;
                    break;

                case 'ㅝ':
                    TouchVirtualKeyBoard('ㅜ');
                    TouchVirtualKeyBoard('ㅓ');
                    return;

                case 'ㅞ':
                    TouchVirtualKeyBoard('ㅜ');
                    TouchVirtualKeyBoard('ㅔ');
                    return;

                case 'ㅟ':
                    TouchVirtualKeyBoard('ㅜ');
                    TouchVirtualKeyBoard('ㅣ');
                    return;

                case 'ㅠ':
                case 'b':
                    x = 903;
                    y = 850;
                    break;

                case 'ㅡ':
                case 'm':
                    x = 1053;
                    y = 856;
                    break;

                case 'ㅢ':
                    TouchVirtualKeyBoard('ㅡ');
                    TouchVirtualKeyBoard('ㅣ');
                    return;

                case 'ㅣ':
                case 'l':
                    x = 1203;
                    y = 769;
                    break;

                case 'ㄳ':
                    TouchVirtualKeyBoard('ㄱ');
                    TouchVirtualKeyBoard('ㅅ');
                    return;

                case 'ㄵ':
                    TouchVirtualKeyBoard('ㄴ');
                    TouchVirtualKeyBoard('ㅈ');
                    return;

                case 'ㄶ':
                    TouchVirtualKeyBoard('ㄴ');
                    TouchVirtualKeyBoard('ㅎ');
                    return;

                case 'ㄺ':
                    TouchVirtualKeyBoard('ㄹ');
                    TouchVirtualKeyBoard('ㄱ');
                    return;

                case 'ㄻ':
                    TouchVirtualKeyBoard('ㄹ');
                    TouchVirtualKeyBoard('ㅁ');
                    return;

                case 'ㄼ':
                    TouchVirtualKeyBoard('ㄹ');
                    TouchVirtualKeyBoard('ㅂ');
                    return;

                case 'ㄽ':
                    TouchVirtualKeyBoard('ㄹ');
                    TouchVirtualKeyBoard('ㅅ');
                    return;

                case 'ㄾ':
                    TouchVirtualKeyBoard('ㄹ');
                    TouchVirtualKeyBoard('ㅌ');
                    return;

                case 'ㄿ':
                    TouchVirtualKeyBoard('ㄹ');
                    TouchVirtualKeyBoard('ㅍ');
                    return;

                case 'ㅀ':
                    TouchVirtualKeyBoard('ㄹ');
                    TouchVirtualKeyBoard('ㅎ');
                    return;

                case 'ㅄ':
                    TouchVirtualKeyBoard('ㅂ');
                    TouchVirtualKeyBoard('ㅅ');
                    return;

                case '1':
                    x = 505;
                    y = 610;
                    break;

                case '2':
                    x = 585;
                    y = 610;
                    break;

                case '3':
                    x = 665;
                    y = 610;
                    break;

                case '4':
                    x = 745;
                    y = 610;
                    break;

                case '5':
                    x = 825;
                    y = 610;
                    break;

                case '6':
                    x = 885;
                    y = 610;
                    break;

                case '7':
                    x = 975;
                    y = 610;
                    break;

                case '8':
                    x = 1060;
                    y = 610;
                    break;

                case '9':
                    x = 1140;
                    y = 610;
                    break;

                case '0':
                    x = 1220;
                    y = 610;
                    break;

                // 인식할 수 없는 문자의 경우 리턴
                default:
                    return;
            }

            // 터치스크린 좌표


            // 쌍자음일 경우 시프트 클릭
            if (isNeedShift)
            {
                // 시프트키 클릭
                AutoItX.MouseClick("LEFT", 459, 857);
                Thread.Sleep(200);

                // 자음 클릭
                AutoItX.MouseClick("LEFT", x, y);
                Thread.Sleep(200);

                // 시프트키 끄기
                AutoItX.MouseClick("LEFT", 459, 857);
            }
            else
            {
                AutoItX.MouseClick("LEFT", x, y);
            }

            // 공용 대기시간
            Thread.Sleep(200);
            LOG.Write("터치스크린 클릭 x = "
                + x.ToString() + ", y = " + y.ToString() + ", SHIFT = " + isNeedShift.ToString());
        }


        #endregion


        // 마우스 클릭 윈도우 API 이벤트
        private void MouseLClick(int x, int y)
        {
            Cursor.Position = new System.Drawing.Point(x, y);
            mouse_event(LBUTTONDOWN, 0, 0, 0, 0);
            Delay(50);
            mouse_event(LBUTTONUP, 0, 0, 0, 0);
        }


        // 마우스 딜레이
        public void Delay(int ms)
        {
            DateTime dateTimeNow = DateTime.Now;
            TimeSpan duration = new TimeSpan(0, 0, 0, 0, ms);
            DateTime dateTimeAdd = dateTimeNow.Add(duration);
            while (dateTimeAdd >= dateTimeNow)
            {
                Application.DoEvents();
                dateTimeNow = DateTime.Now;
            }
            return;
        }


        // 테스트용 버튼 1 클릭 이벤트
        private void button1_Click(object sender, EventArgs e)
        {
            LOG.Write("테스트 시작");
            // 모드선택부터 시작되어야 함.
            DEFAULT_APP_NAME = "Touch Screen";


            // CC 리스트 파일 로딩
            Class.Excel EX = new Class.Excel();            
            DataTable DTccList = EX.ExcelToDataTable(_imagePath + "CCList.xlsx", ".xlsx");


            foreach (DataRow dr in DTccList.Rows)
            {
                LOG.Write("입력된 CC 리스트 : " + dr["CCID"].ToString() + " = " + dr["이름"].ToString());
            }


            try
            {
                // 혹시 모를 화면가림 대비를 위해 golfzonclient에 포커스를 한번 줌
                AutoItX.WinActivate("U1Project");
                Thread.Sleep(1000);


                // 창 활성화            
                AutoItX.WinActivate(DEFAULT_APP_NAME);


                // 어플리케이션 핸들을 받아오지 못하면 강제 종료
                string handle = AutoItX.WinGetHandleAsText(DEFAULT_APP_NAME);
                LOG.Write("handle = " + handle);

                if (handle == "0x00000000")
                {
                    LOG.Write("어플리케이션 창 핸들을 받아오지 못함");
                    throw new Exception("어플리케이션 창 핸들을 받아오지 못함");
                }
                LOG.Write("어플리케이션 창 활성화 성공");


                // 창 활성화까지 대기
                Thread.Sleep(3000);


                int cIndex = 1;
                foreach (DataRow dr in DTccList.Rows)
                {
                    // CCID 및 이름 추출
                    string ccID = dr["CCID"].ToString();
                    string ccName = dr["이름"].ToString();
                    _ccName = ccName;
                    LOG.Write(ccID + ", " + ccName + "시작 // cIndex = " + cIndex.ToString() + "/" + DTccList.Rows.Count.ToString());


                    // 변수 초기화
                    InitParHandiDistance();


                    // 웹데이터 전달
                    GetWebData(ccID);


                    // CC 반복은 여기서부터 시작되어야 함. ★★★★★★★★★★★★★★★
                    // 모드선택 진입 확인 
                    LOG.Write("모드선택 초기상태 확인 시작");
                    bool isModeS = false;
                    int count = 0;
                    while (!isModeS)
                    {
                        // 1분 이상 대기해도 진입못했을 경우 익셉션 드로우 
                        if (count > 30)
                        {
                            throw new Exception("모드선택 진입 실패");
                        }


                        count++;
                        LOG.Write("모드선택 진입 확인 : count = " + count.ToString());
                        Bitmap modeSelect = CapturePartialScreen(150, 250, 150, 50, _tempPath + "모드선택확인");
                        if (ImageToString(ImageUpScale(modeSelect), "kor").Trim() == "스트로크")
                        {
                            isModeS = true;
                            break;
                        }                        
                        Thread.Sleep(2000);                        
                    }


                    // 다음버튼 클릭 --> CC 선택 이동
                    MouseLClick(1770, 1020);
                    LOG.Write("다음 버튼 클릭");
                    Thread.Sleep(5000);


                    // 검색버튼 클릭
                    LOG.Write("검색 버튼 클릭");
                    MouseLClick(1686, 88);
                    Thread.Sleep(3000);


                    // CC선택 진입 확인
                    LOG.Write("CC선택 CC입력 상태 확인 시작");
                    Bitmap ccSelect = CapturePartialScreen(1010, 90, 170, 70, _tempPath + "CC선택확인");
                    if (ImageToString(ccSelect, "kor").Trim() != "주십시오.")
                    {
                        //throw new Exception("CC선택 진입 실패");
                    }


                    // CC명 입력
                    LOG.Write("CC명 입력");
                    AutoItX.Send(ccName);
                    Thread.Sleep(3000);


                    // ★★★★★★★★★★★★★ 보완 필요 (엔터키 안먹을때 있음)
                    LOG.Write("ENTER 클릭");
                    MouseLClick(1560, 834);
                    Thread.Sleep(3000);


                    // 검색된 CC 클릭
                    LOG.Write("검색된 첫 CC 클릭");
                    MouseLClick(283, 523);
                    Thread.Sleep(3000);


                    // 다음버튼 클릭 -> 라운드 설정 이동
                    LOG.Write("라운드 설정 이동");
                    MouseLClick(1770, 1020);
                    Thread.Sleep(5000);


                    // 라운드 설정 진입 확인
                    LOG.Write("라운드 설정 진입 확인");
                    Bitmap roundSetting = CapturePartialScreen(137, 168, 151, 53, _tempPath + "라운드설정확인");
                    if (ImageToString(roundSetting, "kor").Trim() != "라운드 방식")
                    {
                        //throw new Exception("라운드 설정 진입 실패");
                    }


                    // 다음버튼 클릭 -> 개인 설정
                    LOG.Write("개인 설정 이동");
                    MouseLClick(1770, 1020);
                    Thread.Sleep(5000);


                    // 다음버튼 클릭 -> 라운드 시작 1홀 이동
                    LOG.Write("라운드 시작");
                    MouseLClick(1770, 1020);


                    // 라운드 시작 30초 대기 (1홀 동영상 광고)
                    Thread.Sleep(30000);


                    // 홀스킵 하면서 티별 길이 데이터 추출
                    for (int h = 1; h < 19; h++)
                    {
                        bool isHoleStart = false;
                        count = 0;

                        // 2초 단위로 홀 시작지점인지 판별함
                        while (!isHoleStart)
                        {
                            // 2분 이상 대기할 경우 익셉션 드로우
                            if (count > 60)
                            {
                                throw new Exception(h.ToString() + "홀 진입 실패");
                            }

                            count++;
                            LOG.Write("홀시작 판정 시작 / count = " + count.ToString());
                            Bitmap hole = CapturePartialScreen(70, 70, 70, 55, _tempPath + "hole" + h.ToString());
                            if (ImageToString(hole, "eng").Trim() == "Hole")
                            {
                                isHoleStart = true;
                                LOG.Write("isHoleStart = " + isHoleStart.ToString());
                                break;
                            }

                            Thread.Sleep(2000);
                        }
                        LOG.Write("로딩 종료, " + h.ToString() + "홀 시작됨");

                        // 1홀일 때만 로그에서 기준파/핸디캡 정보 가져옴
                        if (h == 1)
                        {
                            GetParAndHandicapToLog(ccID);
                        }


                        // 홀번호 읽고 저장
                        Bitmap holeNum = CapturePartialScreen(140, 75, 70, 55, _tempPath + "h" + h.ToString());
                        LOG.Write("홀번호 읽음 = " + ImageToString(ImageUpScale(holeNum), "eng").Trim());
                        Thread.Sleep(1000);
                        

                        // 로그에서 티 거리 정보 가져옴
                        GetTeeDistanceToLog(ccID, h - 1);


                        // 다음 메뉴 호출
                        MouseLClick(1826, 537);
                        Thread.Sleep(1000);


                        // 홀 넘기기 버튼
                        MouseLClick(1642, 444);
                        Thread.Sleep(1000);


                        // 확인 버튼 클릭
                        MouseLClick(1096, 644);
                        Thread.Sleep(15000);
                    }


                    // 결과를 템프에 기록
                    WriteTempResult(cIndex);


                    // 라운드 종료 버튼 클릭
                    MouseLClick(1777, 1020);
                    Thread.Sleep(1000);


                    // 유지 종료 버튼 클릭
                    MouseLClick(962, 723);
                    Thread.Sleep(15000);


                    cIndex++;
                }
            }
            catch (Exception ex)
            {
                LOG.Write(ex.Message);                
            }
            finally
            {
                EX.Excel2007Export(_DTresult, _resultPath + "테스트 결과_전체.xlsx");
                AutoItX.WinActivate(Application.ProductName);
                MessageBox.Show("작업완료!!!");
            }
        }


        // 게임로그 가져오기
        private string[] GetGameLog()
        {
            string path = @"C:\Users\Golfzon\Desktop\2023_03_09\GolfzonVision2NX\U1Project\Saved\Logs\U1Project.log";
            //string path2 = @"C:\Users\ozsec\Downloads\U1Project.log"; <-- 테스트용

            var list = new List<string>();


            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var sr = new StreamReader(fs, Encoding.Default))
            {
                string line;

                while ((line = sr.ReadLine()) != null)
                {
                    list.Add(line);
                }
            }
            string[] logs = list.ToArray();


            return logs;
        }


        private string[] GetTeeDistance(string raw, string firstTee, string secondTee)
        {
            string temp = raw.Substring(raw.IndexOf(firstTee) + firstTee.Length, raw.IndexOf(secondTee) - raw.IndexOf(firstTee) - firstTee.Length);
            string[] teeDistance = temp.Split('m');
            Array.Resize(ref teeDistance, 9);
            

            return teeDistance;
        }


        // 게임로그에서 기준파와 핸디캡을 불러오는 로직 (_par, _handi)
        private void GetParAndHandicapToLog(string ccID)
        {
            // 로그파일을 읽어와서 전장거리 추출
            string[] gameLog = GetGameLog();
            int index = 0;            


            for (int l = gameLog.Length - 1; l >= 0; l--)
            {                
                try
                {
                    string[] temp = gameLog[l].Split(']');
                    string log = temp[2];


                    // 파, 핸디캡 정보 검색
                    if (log.Contains("[CC INFORMATION"))
                    {
                        LOG.Write("Par/Handi Search Line = " + l.ToString("N0"));
                        string[] ccIDTemp = gameLog[l].Split('|');
                        string ccID2 = ccIDTemp[1].Substring(8).Trim();


                        // CCID 대조 후 같을 때만 
                        if (ccID == ccID2)
                        {
                            // 로그에서 파/핸디캡 추출
                            for (int i = l + 3; i < l + 21; i++)
                            {
                                string[] parTemp = gameLog[i].Split('|');
                                //_par[index] = int.Parse(parTemp[1].Trim());
                                //_handi[index] = int.Parse(parTemp[2].Trim());

                                int webPar = 0;
                                int.TryParse(_DTtempResult.Rows[index]["WebPar"].ToString(), out webPar);
                                int cliPar = int.Parse(parTemp[1].Trim());
                                int handi = int.Parse(parTemp[2].Trim());


                                _DTtempResult.Rows[index]["CliPar"] = cliPar;
                                _DTtempResult.Rows[index]["Handicap"] = handi;


                                LOG.Write("index = " + index.ToString()
                                            + " / WebPar = " + webPar.ToString()
                                            + " / CliPar = " + cliPar.ToString()
                                            + " / Handicap =" + handi.ToString());


                                // P/F 판별
                                if (webPar == cliPar)
                                {
                                    _DTtempResult.Rows[index]["Result"] = "Pass";
                                }
                                else
                                {
                                    _DTtempResult.Rows[index]["Result"] = "Fail";
                                }


                                index += 6;
                            }
                        }

                        // 데이터를 찾으면 추가 읽기를 중단함
                        break;
                    }
                }
                catch (Exception ex)
                {
                    LOG.Write(ex.Message);
                }
            }
        }


        // 게임로그에서 티별 거리를 불러오는 로직 (_distance) -> 홀마다 호출
        private void GetTeeDistanceToLog(string ccID, int hIndex)
        {
            string[] gameLog = GetGameLog();
            int tIndex = 0;
            

            for (int l = gameLog.Length - 1; l >= 0; l--)
            {                
                try
                {
                    string[] temp = gameLog[l].Split(']');
                    string log = temp[2];


                    // 티 별 길이 정보 검색 (매 홀 시작마다)
                    if (log.Contains("[HOLE INFORMATION"))
                    {
                        LOG.Write("Tee Distance Search Line = " + l.ToString("N0"));
                        string[] ccIDTemp = gameLog[l].Split('|');
                        string ccID2 = ccIDTemp[1].Substring(8).Trim();


                        // CCID 대조 후 같을 때만 
                        if (ccID == ccID2)
                        {
                            for (int i = l + 5; i < l + 20; i = i + 3)
                            {
                                string[] disTemp = gameLog[i].Split('|');
                                int dis = int.Parse(disTemp[1].Trim().Substring(14));
                                //_distance[tIndex, hIndex] = dis;                                
                                LOG.Write(disTemp[9].Trim() + " / Distance = " + dis.ToString());


                                // 홀 정보 데이터가 기록된 행 기준 (0, 7, 13, 19 ...)
                                int index = hIndex * 6 + 1;


                                // 홀 및 웹파 추출 -> 필요없는 부분
                                int hole = Convert.ToInt16(_DTtempResult.Rows[index]["Hole"]);
                                int webPar = 0;
                                int.TryParse(_DTtempResult.Rows[index]["WebPar"].ToString(), out webPar);


                                // 클라 티간 거리 입력
                                _DTtempResult.Rows[index + tIndex]["CliDistance"] = dis;
                                int webDis = 0;
                                int.TryParse(_DTtempResult.Rows[index + tIndex]["WebDistance"].ToString(), out webDis);
                                LOG.Write("WebPar = " + webPar.ToString() + " / WebDis = " + webDis.ToString() + " / CliDis = " + dis.ToString());


                                // P/F 판별
                                if (webDis == dis)
                                {
                                    _DTtempResult.Rows[index + tIndex]["Result"] = "Pass";
                                }
                                else
                                {
                                    _DTtempResult.Rows[index + tIndex]["Result"] = "Fail";
                                }


                                tIndex++;

                                /*
                                // tempResult에서 티 별 거리를 대조하여 다르면 스크린 샷                                
                                for (int r = 0; r < _DTtempResult.Rows.Count; r++)
                                {   
                                    if (hole == hIndex && webPar != 0)
                                    {                                        
                                        _DTtempResult.Rows[r + 1 + tIndex]["CliDistance"] = dis;
                                        int webDis = 0;
                                        int.TryParse(_DTtempResult.Rows[r]["WebDistance"].ToString(), out webDis);
                                        LOG.Write("WebPar = " + webPar.ToString() + " / WebDis = " + webDis.ToString() + " / CliDis = " + dis.ToString());


                                        
                                    }
                                }                                
                                */
                            }


                            tIndex = 0;
                        }


                        // 한번 데이터를 찾으면 검색 중단
                        break;
                    }
                }
                catch (Exception ex)
                {
                    LOG.Write(ex.Message);
                }
            }
        }


        // 저장된 기준파/핸디캡/티별거리를 임시 결과 테이블에 씀 
        private void WriteTempResult(int cIndex)
        {
            Class.Excel EX = new Class.Excel(); 
            int hIndex = -1;
            int tIndex = 0;
            
            /*
            // 찾은 데이터 기록 및 정상여부 판별
            foreach (DataRow dr in _DTtempResult.Rows)
            {
                int webPar = 0;
                int.TryParse(dr["WebPar"].ToString(), out webPar);


                // 테이블에 기준파 값이 있으면 
                if (webPar > 0 || tIndex > 4)
                {
                    hIndex++;
                    tIndex = 0;


                    dr["CliPar"] = _par[hIndex];
                    dr["Handicap"] = _handi[hIndex];


                    // 기준파 값이 같으면 Pass, 다르면 Fail 기록
                    LOG.Write("CliPar = " + _par[hIndex].ToString() + " / WebPar = " + webPar.ToString() + " / Handicap = " + _handi[hIndex].ToString());
                    if (_par[hIndex] == webPar && _handi[hIndex] != 0)
                    {
                        dr["Result"] = "Pass";
                    }
                    else
                    {
                        dr["Result"] = "Fail";
                    }
                }
                // 기준파 값이 없으면 전장거리 입력
                else
                {
                    dr["CliDistance"] = _distance[tIndex, hIndex];
                    float webDist = float.Parse(dr["WebDistance"].ToString());
                    LOG.Write("CliDist = " + _distance[tIndex, hIndex].ToString() + " / WebDist = " + webDist.ToString());


                    if (webDist == _distance[tIndex, hIndex])
                    {
                        dr["Result"] = "Pass";
                    }
                    else
                    {
                        dr["Result"] = "Fail";
                    }


                    tIndex++;                   
                }
            }
            */

            // 임시 결과 엑셀 익스포트
            EX.Excel2007Export(_DTtempResult, _tempPath + "TempResult.xlsx");


            // 결과에 임시결과 머지 후 익스포트
            _DTresult.Merge(_DTtempResult);
            LOG.Write("테이블 머지 결과 = " + _DTresult.Rows.Count.ToString() + "Rows");
            EX.Excel2007Export(_DTresult, _resultPath + "테스트 결과_" + _ccName.ToString() + ".xlsx");


            // 임시 결과 테이블 내용삭제
            _DTtempResult.Clear();
        }


        // 파/핸디캡/티별거리 변수 초기화
        private void InitParHandiDistance()
        {
            // _par, _handi 초기화
            for (int i = 0; i < 18; i++)
            {
                _par[i] = 0;
                _handi[i] = 0;


                for (int j = 0; j < 5; j++)
                {
                    _distance[j, i] = 0;
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            Class.Excel EX = new Class.Excel();
            GetWebData("102294784");            
            _DTresult.Merge(_DTtempResult);
            EX.Excel2007Export(_DTresult, _tempPath + "Result.xlsx");

            GetWebData("100000877");
            _DTresult.Merge(_DTtempResult);
            EX.Excel2007Export(_DTresult, _tempPath + "Result.xlsx");





            /*
            LOG.Write("Recording Start");

            _recordingStartTime = DateTime.Now;

            // 레코드 스레딩
            _tRecording = new Thread(ThreadRecording);
            _tRecording.Start();
            */
        }
        

        private void button3_Click(object sender, EventArgs e)
        {
            DEFAULT_APP_NAME = "Touch Screen";
            AutoItX.WinActivate(DEFAULT_APP_NAME);
            Thread.Sleep(1000);


            string[] temp = TXTtest3.Text.Split(',');
            int x = int.Parse(temp[0]);
            int y = int.Parse(temp[1]);
            int width = int.Parse(temp[2]);
            int height = int.Parse(temp[3]);


            CapturePartialScreen(x, y, width, height, _imagePath + "h");


            AutoItX.WinActivate(Application.ProductName);

            /*
            _tRecording.Abort();
            _tRecording.Join();

            TimeSpan recTime = DateTime.Now - _recordingStartTime;
            double term = recTime.TotalMilliseconds;

            // 생성된 이미지 파일 수로 영상의 fps를 계산
            double fps = (double)_fileCount / (term / 1000);
            fps = Math.Round(fps);

            LOG.Write("_fileCount = " + _fileCount.ToString("N0")
                + " ,term = " + term.ToString("N0")
                + "ms ,fps = " + fps.ToString());

            if (fps != 0)
            {
                EncodingVideo((int)fps, "Thread Test");
            }
            else
            {
                LOG.Write("계산된 fps가 0이어서 영상을 인코딩할 수 없음.");
            }
            */

        }








    }
}

