# =============================================================
#  forms.ps1 - UserForm 레이아웃 정의
#
#  .frm 파일은 .frx 바이너리 짝이 필요해 손으로 만들면 깨지기 쉽다.
#  그래서 폼은 파일로 두지 않고 build.ps1 이 VBComponents.Add(3) 로 만든 뒤
#  Designer.Controls.Add 로 컨트롤을 코드로 배치한다.
#
#  각 컨트롤:  T=종류  N=이름  L/Y/W/H=위치·크기  C=Caption  P=추가 속성
# =============================================================

$Global:FormDefs = @(

  #-----------------------------------------------------------
  @{ Name='frmTask'; Caption='업무 등록'; Width=430; Height=310; Controls=@(
      @{T='Label';         N='lbl1';       L=10;  Y=12;  W=60;  H=18; C='제목'}
      @{T='TextBox';       N='txtTitle';   L=80;  Y=10;  W=330; H=20}
      @{T='Label';         N='lbl2';       L=10;  Y=40;  W=60;  H=18; C='날짜'}
      @{T='TextBox';       N='txtDate';    L=80;  Y=38;  W=100; H=20}
      @{T='Label';         N='lbl3';       L=195; Y=40;  W=80;  H=18; C='시간(선택)'}
      @{T='TextBox';       N='txtTime';    L=280; Y=38;  W=80;  H=20}
      @{T='Label';         N='lbl4';       L=10;  Y=68;  W=60;  H=18; C='내용'}
      @{T='TextBox';       N='txtContent'; L=80;  Y=66;  W=330; H=160; P=@{MultiLine=$true; ScrollBars=2; EnterKeyBehavior=$true; WordWrap=$true}}
      @{T='CommandButton'; N='cmdOK';      L=80;  Y=238; W=90;  H=28; C='저장'}
      @{T='CommandButton'; N='cmdCancel';  L=180; Y=238; W=90;  H=28; C='취소'}
      @{T='CommandButton'; N='cmdDelete';  L=320; Y=238; W=90;  H=28; C='삭제'}
  )}

  #-----------------------------------------------------------
  @{ Name='frmContact'; Caption='연락하기 등록'; Width=440; Height=215; Controls=@(
      @{T='Label';         N='lbl1';      L=10;  Y=12;  W=60;  H=18; C='이름'}
      @{T='TextBox';       N='txtName';   L=80;  Y=10;  W=130; H=20}
      @{T='CommandButton'; N='cmdFind';   L=218; Y=9;   W=110; H=22; C='명부에서 찾기'}
      @{T='Label';         N='lbl2';      L=10;  Y=40;  W=60;  H=18; C='부서'}
      @{T='TextBox';       N='txtDept';   L=80;  Y=38;  W=130; H=20}
      @{T='Label';         N='lblExt';    L=218; Y=40;  W=200; H=18; C=''}
      @{T='Label';         N='lbl3';      L=10;  Y=68;  W=110; H=18; C='내선 직접입력'}
      @{T='TextBox';       N='txtExt';    L=125; Y=66;  W=85;  H=20}
      @{T='Label';         N='lbl4';      L=10;  Y=96;  W=60;  H=18; C='날짜(선택)'}
      @{T='TextBox';       N='txtDate';   L=80;  Y=94;  W=100; H=20}
      @{T='Label';         N='lbl5';      L=195; Y=96;  W=80;  H=18; C='시간(선택)'}
      @{T='TextBox';       N='txtTime';   L=280; Y=94;  W=80;  H=20}
      @{T='CommandButton'; N='cmdOK';     L=80;  Y=130; W=90;  H=28; C='저장'}
      @{T='CommandButton'; N='cmdCancel'; L=180; Y=130; W=90;  H=28; C='취소'}
      @{T='CommandButton'; N='cmdDelete'; L=320; Y=130; W=90;  H=28; C='삭제'}
  )}

  #-----------------------------------------------------------
  @{ Name='frmShare'; Caption='업무공유 등록'; Width=480; Height=345; Controls=@(
      @{T='Label';         N='lbl1';       L=10;  Y=12;  W=70;  H=18; C='문서번호'}
      @{T='TextBox';       N='txtDocNo';   L=90;  Y=10;  W=150; H=20}
      @{T='Label';         N='lblWarn';    L=250; Y=12;  W=210; H=18; C=''}
      @{T='Label';         N='lbl2';       L=10;  Y=40;  W=70;  H=18; C='제목'}
      @{T='TextBox';       N='txtTitle';   L=90;  Y=38;  W=370; H=20}
      @{T='Label';         N='lbl3';       L=10;  Y=68;  W=70;  H=18; C='등록일'}
      @{T='TextBox';       N='txtReg';     L=90;  Y=66;  W=100; H=20}
      @{T='Label';         N='lbl4';       L=205; Y=68;  W=80;  H=18; C='회신기한'}
      @{T='TextBox';       N='txtDue';     L=290; Y=66;  W=100; H=20}
      @{T='Label';         N='lbl5';       L=10;  Y=96;  W=70;  H=18; C='내용'}
      @{T='TextBox';       N='txtContent'; L=90;  Y=94;  W=370; H=160; P=@{MultiLine=$true; ScrollBars=2; EnterKeyBehavior=$true; WordWrap=$true}}
      @{T='CommandButton'; N='cmdOK';      L=90;  Y=268; W=90;  H=28; C='저장'}
      @{T='CommandButton'; N='cmdCancel';  L=190; Y=268; W=90;  H=28; C='취소'}
      @{T='CommandButton'; N='cmdDelete';  L=370; Y=268; W=90;  H=28; C='삭제'}
  )}

  #-----------------------------------------------------------
  @{ Name='frmLunch'; Caption='점심약속 등록'; Width=440; Height=210; Controls=@(
      @{T='Label';         N='lbl1';      L=10;  Y=12;  W=70;  H=18; C='대상 이름'}
      @{T='TextBox';       N='txtTarget'; L=90;  Y=10;  W=150; H=20}
      @{T='Label';         N='lbl2';      L=10;  Y=40;  W=70;  H=18; C='날짜'}
      @{T='TextBox';       N='txtDate';   L=90;  Y=38;  W=100; H=20}
      @{T='Label';         N='lbl3';      L=205; Y=40;  W=80;  H=18; C='시간(선택)'}
      @{T='TextBox';       N='txtTime';   L=290; Y=38;  W=80;  H=20}
      @{T='Label';         N='lbl4';      L=10;  Y=68;  W=70;  H=18; C='장소'}
      @{T='TextBox';       N='txtPlace';  L=90;  Y=66;  W=280; H=20}
      @{T='Label';         N='lblInfo';   L=90;  Y=94;  W=330; H=18; C=''}
      @{T='CommandButton'; N='cmdOK';     L=90;  Y=125; W=90;  H=28; C='저장'}
      @{T='CommandButton'; N='cmdCancel'; L=190; Y=125; W=90;  H=28; C='취소'}
      @{T='CommandButton'; N='cmdDelete'; L=320; Y=125; W=90;  H=28; C='삭제'}
  )}

  #-----------------------------------------------------------
  @{ Name='frmPicker'; Caption='동명이인 선택'; Width=490; Height=305; Controls=@(
      @{T='Label';         N='lblInfo';   L=10;  Y=10;  W=460; H=18; C=''}
      @{T='ListBox';       N='lstPeople'; L=10;  Y=34;  W=460; H=190; P=@{ColumnCount=4; ColumnWidths='110 pt;130 pt;70 pt;90 pt'; ColumnHeads=$false}}
      @{T='CommandButton'; N='cmdOK';     L=275; Y=234; W=90;  H=28; C='선택'}
      @{T='CommandButton'; N='cmdCancel'; L=375; Y=234; W=90;  H=28; C='취소'}
  )}

  #-----------------------------------------------------------
  @{ Name='frmDetail'; Caption='내용 보기'; Width=530; Height=415; Controls=@(
      @{T='Label';         N='lblTitle'; L=10;  Y=10;  W=500; H=20; C=''}
      @{T='TextBox';       N='txtBody';  L=10;  Y=36;  W=500; H=300; P=@{MultiLine=$true; ScrollBars=3; WordWrap=$true; Locked=$true}}
      @{T='CommandButton'; N='cmdCopy';  L=310; Y=345; W=100; H=28; C='내용 복사'}
      @{T='CommandButton'; N='cmdClose'; L=420; Y=345; W=90;  H=28; C='닫기'}
  )}

  #-----------------------------------------------------------
  @{ Name='frmExport'; Caption='업무공유 내보내기'; Width=450; Height=245; Controls=@(
      @{T='Label';         N='lbl1';      L=10;  Y=10;  W=200; H=18; C='내보낼 범위를 고르세요'}
      @{T='OptionButton';  N='optAll';    L=20;  Y=32;  W=150; H=18; C='전체'}
      @{T='OptionButton';  N='optMine';   L=20;  Y=54;  W=180; H=18; C='내가 등록한 것만'}
      @{T='OptionButton';  N='optRecent'; L=20;  Y=76;  W=90;  H=18; C='최근 N일'}
      @{T='TextBox';       N='txtDays';   L=115; Y=74;  W=40;  H=20}
      @{T='Label';         N='lbl2';      L=160; Y=76;  W=20;  H=18; C='일'}
      @{T='OptionButton';  N='optOpen';   L=20;  Y=98;  W=200; H=18; C='회신기한 미도래분만'}
      @{T='Label';         N='lblInfo';   L=20;  Y=128; W=400; H=32; C=''; P=@{WordWrap=$true}}
      @{T='CommandButton'; N='cmdCopy';   L=20;  Y=168; W=200; H=32; C='복사'}
      @{T='CommandButton'; N='cmdClose';  L=330; Y=168; W=90;  H=32; C='닫기'}
  )}

  #-----------------------------------------------------------
  @{ Name='frmImport'; Caption='업무공유 가져오기'; Width=570; Height=450; Controls=@(
      @{T='Label';         N='lbl1';       L=10;  Y=10;  W=400; H=18; C='받은 텍스트를 아래 칸에 붙여넣으세요 (Ctrl+V)'}
      @{T='TextBox';       N='txtPaste';   L=10;  Y=32;  W=540; H=100; P=@{MultiLine=$true; ScrollBars=3; WordWrap=$false}}
      @{T='Label';         N='lblStatus';  L=10;  Y=140; W=540; H=32; C=''; P=@{WordWrap=$true}}
      @{T='ListBox';       N='lstPreview'; L=10;  Y=178; W=540; H=170; P=@{ColumnCount=3; ColumnWidths='70 pt;130 pt;330 pt'; ColumnHeads=$false}}
      @{T='CommandButton'; N='cmdApply';   L=10;  Y=358; W=120; H=32; C='반영'}
      @{T='CommandButton'; N='cmdReset';   L=140; Y=358; W=140; H=32; C='받은 조각 버리기'}
      @{T='CommandButton'; N='cmdClose';   L=460; Y=358; W=90;  H=32; C='닫기'}
  )}
)
