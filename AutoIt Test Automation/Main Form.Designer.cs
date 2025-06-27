namespace Golfzon_TVNX_Test
{
    partial class FORMmain
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.BTNstart = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.TXTlog = new System.Windows.Forms.RichTextBox();
            this.TXTpassword = new System.Windows.Forms.TextBox();
            this.LBLpassword = new System.Windows.Forms.Label();
            this.CBBgsType = new System.Windows.Forms.ComboBox();
            this.LBLgsType = new System.Windows.Forms.Label();
            this.TXTccName = new System.Windows.Forms.TextBox();
            this.LBLccName = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.TXThole = new System.Windows.Forms.TextBox();
            this.LBLhole = new System.Windows.Forms.Label();
            this.LBLrepeat = new System.Windows.Forms.Label();
            this.TXTrepeat = new System.Windows.Forms.TextBox();
            this.LBLbuildVer = new System.Windows.Forms.Label();
            this.CHKnasmo = new System.Windows.Forms.CheckBox();
            this.LBLholeEx = new System.Windows.Forms.Label();
            this.CHKusePassword = new System.Windows.Forms.CheckBox();
            this.TXTclientFolderName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // BTNstart
            // 
            this.BTNstart.Font = new System.Drawing.Font("맑은 고딕", 11.25F);
            this.BTNstart.Location = new System.Drawing.Point(370, 33);
            this.BTNstart.Margin = new System.Windows.Forms.Padding(4);
            this.BTNstart.Name = "BTNstart";
            this.BTNstart.Size = new System.Drawing.Size(163, 37);
            this.BTNstart.TabIndex = 0;
            this.BTNstart.Text = "시작";
            this.BTNstart.UseVisualStyleBackColor = true;
            this.BTNstart.Click += new System.EventHandler(this.BTNstart_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("맑은 고딕", 11.25F);
            this.button1.Location = new System.Drawing.Point(442, 26);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(129, 30);
            this.button1.TabIndex = 4;
            this.button1.Text = "Test1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("맑은 고딕", 11.25F);
            this.button2.Location = new System.Drawing.Point(442, 62);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(129, 30);
            this.button2.TabIndex = 5;
            this.button2.Text = "Test2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // TXTlog
            // 
            this.TXTlog.Font = new System.Drawing.Font("나눔명조", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.TXTlog.Location = new System.Drawing.Point(22, 361);
            this.TXTlog.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.TXTlog.Name = "TXTlog";
            this.TXTlog.ReadOnly = true;
            this.TXTlog.Size = new System.Drawing.Size(983, 253);
            this.TXTlog.TabIndex = 3;
            this.TXTlog.Text = "";
            // 
            // TXTpassword
            // 
            this.TXTpassword.Font = new System.Drawing.Font("맑은 고딕", 11.25F);
            this.TXTpassword.Location = new System.Drawing.Point(114, 33);
            this.TXTpassword.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.TXTpassword.Name = "TXTpassword";
            this.TXTpassword.Size = new System.Drawing.Size(200, 27);
            this.TXTpassword.TabIndex = 6;
            this.TXTpassword.Text = "1";
            // 
            // LBLpassword
            // 
            this.LBLpassword.AutoSize = true;
            this.LBLpassword.Location = new System.Drawing.Point(18, 36);
            this.LBLpassword.Name = "LBLpassword";
            this.LBLpassword.Size = new System.Drawing.Size(93, 20);
            this.LBLpassword.TabIndex = 7;
            this.LBLpassword.Text = "GS 비밀번호";
            // 
            // CBBgsType
            // 
            this.CBBgsType.FormattingEnabled = true;
            this.CBBgsType.Items.AddRange(new object[] {
            "REAL",
            "VISION",
            "TWOVISION"});
            this.CBBgsType.Location = new System.Drawing.Point(114, 79);
            this.CBBgsType.Name = "CBBgsType";
            this.CBBgsType.Size = new System.Drawing.Size(200, 28);
            this.CBBgsType.TabIndex = 8;
            // 
            // LBLgsType
            // 
            this.LBLgsType.AutoSize = true;
            this.LBLgsType.Location = new System.Drawing.Point(19, 82);
            this.LBLgsType.Name = "LBLgsType";
            this.LBLgsType.Size = new System.Drawing.Size(63, 20);
            this.LBLgsType.TabIndex = 7;
            this.LBLgsType.Text = "GS 타입";
            // 
            // TXTccName
            // 
            this.TXTccName.Font = new System.Drawing.Font("맑은 고딕", 11.25F);
            this.TXTccName.Location = new System.Drawing.Point(114, 127);
            this.TXTccName.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.TXTccName.Name = "TXTccName";
            this.TXTccName.Size = new System.Drawing.Size(200, 27);
            this.TXTccName.TabIndex = 6;
            this.TXTccName.Text = "gtour 마운틴";
            // 
            // LBLccName
            // 
            this.LBLccName.AutoSize = true;
            this.LBLccName.Location = new System.Drawing.Point(19, 130);
            this.LBLccName.Name = "LBLccName";
            this.LBLccName.Size = new System.Drawing.Size(44, 20);
            this.LBLccName.TabIndex = 7;
            this.LBLccName.Text = "CC명";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(442, 98);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(129, 30);
            this.button3.TabIndex = 9;
            this.button3.Text = "Test3";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // TXThole
            // 
            this.TXThole.Font = new System.Drawing.Font("맑은 고딕", 11.25F);
            this.TXThole.Location = new System.Drawing.Point(114, 178);
            this.TXThole.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.TXThole.Name = "TXThole";
            this.TXThole.Size = new System.Drawing.Size(200, 27);
            this.TXThole.TabIndex = 11;
            this.TXThole.Text = "1";
            // 
            // LBLhole
            // 
            this.LBLhole.AutoSize = true;
            this.LBLhole.Location = new System.Drawing.Point(19, 181);
            this.LBLhole.Name = "LBLhole";
            this.LBLhole.Size = new System.Drawing.Size(54, 20);
            this.LBLhole.TabIndex = 12;
            this.LBLhole.Text = "작업홀";
            // 
            // LBLrepeat
            // 
            this.LBLrepeat.AutoSize = true;
            this.LBLrepeat.Location = new System.Drawing.Point(19, 231);
            this.LBLrepeat.Name = "LBLrepeat";
            this.LBLrepeat.Size = new System.Drawing.Size(69, 20);
            this.LBLrepeat.TabIndex = 13;
            this.LBLrepeat.Text = "반복횟수";
            // 
            // TXTrepeat
            // 
            this.TXTrepeat.Font = new System.Drawing.Font("맑은 고딕", 11.25F);
            this.TXTrepeat.Location = new System.Drawing.Point(114, 228);
            this.TXTrepeat.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.TXTrepeat.Name = "TXTrepeat";
            this.TXTrepeat.Size = new System.Drawing.Size(200, 27);
            this.TXTrepeat.TabIndex = 11;
            this.TXTrepeat.Text = "1";
            // 
            // LBLbuildVer
            // 
            this.LBLbuildVer.AutoSize = true;
            this.LBLbuildVer.Location = new System.Drawing.Point(871, 9);
            this.LBLbuildVer.Name = "LBLbuildVer";
            this.LBLbuildVer.Size = new System.Drawing.Size(133, 20);
            this.LBLbuildVer.TabIndex = 7;
            this.LBLbuildVer.Text = "ver. 000000000000";
            this.LBLbuildVer.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // CHKnasmo
            // 
            this.CHKnasmo.AutoSize = true;
            this.CHKnasmo.Location = new System.Drawing.Point(114, 276);
            this.CHKnasmo.Name = "CHKnasmo";
            this.CHKnasmo.Size = new System.Drawing.Size(143, 24);
            this.CHKnasmo.TabIndex = 14;
            this.CHKnasmo.Text = "나스모 광고 체크";
            this.CHKnasmo.UseVisualStyleBackColor = true;
            // 
            // LBLholeEx
            // 
            this.LBLholeEx.AutoSize = true;
            this.LBLholeEx.Location = new System.Drawing.Point(321, 181);
            this.LBLholeEx.Name = "LBLholeEx";
            this.LBLholeEx.Size = new System.Drawing.Size(96, 20);
            this.LBLholeEx.TabIndex = 15;
            this.LBLholeEx.Text = "(ex. 1, 3, 5-8)";
            // 
            // CHKusePassword
            // 
            this.CHKusePassword.AutoSize = true;
            this.CHKusePassword.Checked = true;
            this.CHKusePassword.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CHKusePassword.Location = new System.Drawing.Point(114, 306);
            this.CHKusePassword.Name = "CHKusePassword";
            this.CHKusePassword.Size = new System.Drawing.Size(173, 24);
            this.CHKusePassword.TabIndex = 16;
            this.CHKusePassword.Text = "관리자 패스워드 사용";
            this.CHKusePassword.UseVisualStyleBackColor = true;
            // 
            // TXTclientFolderName
            // 
            this.TXTclientFolderName.Location = new System.Drawing.Point(121, 28);
            this.TXTclientFolderName.Name = "TXTclientFolderName";
            this.TXTclientFolderName.Size = new System.Drawing.Size(314, 27);
            this.TXTclientFolderName.TabIndex = 17;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 20);
            this.label1.TabIndex = 18;
            this.label1.Text = "클라 폴더 이름";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.TXTclientFolderName);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Location = new System.Drawing.Point(407, 204);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(596, 143);
            this.groupBox1.TabIndex = 19;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "NX CC 검증";
            // 
            // FORMmain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(1017, 637);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.CHKusePassword);
            this.Controls.Add(this.LBLholeEx);
            this.Controls.Add(this.CHKnasmo);
            this.Controls.Add(this.LBLrepeat);
            this.Controls.Add(this.LBLhole);
            this.Controls.Add(this.TXTrepeat);
            this.Controls.Add(this.TXThole);
            this.Controls.Add(this.CBBgsType);
            this.Controls.Add(this.LBLccName);
            this.Controls.Add(this.LBLgsType);
            this.Controls.Add(this.LBLbuildVer);
            this.Controls.Add(this.LBLpassword);
            this.Controls.Add(this.TXTccName);
            this.Controls.Add(this.TXTpassword);
            this.Controls.Add(this.TXTlog);
            this.Controls.Add(this.BTNstart);
            this.Font = new System.Drawing.Font("맑은 고딕", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FORMmain";
            this.Text = "Golfzon TVNX Test";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BTNstart;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox TXTpassword;
        private System.Windows.Forms.Label LBLpassword;
        private System.Windows.Forms.ComboBox CBBgsType;
        private System.Windows.Forms.Label LBLgsType;
        private System.Windows.Forms.TextBox TXTccName;
        private System.Windows.Forms.Label LBLccName;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox TXThole;
        private System.Windows.Forms.Label LBLhole;
        private System.Windows.Forms.Label LBLrepeat;
        private System.Windows.Forms.TextBox TXTrepeat;
        private System.Windows.Forms.Label LBLbuildVer;
        private System.Windows.Forms.CheckBox CHKnasmo;
        private System.Windows.Forms.Label LBLholeEx;
        public System.Windows.Forms.RichTextBox TXTlog;
        private System.Windows.Forms.CheckBox CHKusePassword;
        private System.Windows.Forms.TextBox TXTclientFolderName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}

