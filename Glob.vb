Imports Microsoft.VisualBasic

Public Class Glob
    Public Const TIME_11_59_PM As String = "23:59:59.999"

    Public Enum GUIColor
        BG_DARK = 1
        BG_LIGHT = 2
        FG_TITLE = 3

        ORANGE_BRIGHT = 4
        ORANGE_LIGHT = 5
        YELLOW_WHITE = 6
        YELLOW_BRIGHT = 7
        GREY = 8

        OLIVE_DARK = 10
        OLIVE_LIGHT = 11
        CREAM = 12
    End Enum

    ' MONTHLY: monthly billing periods, using the system to generate bills, input payments and track balances
    ' ADHOC: adhoc billing periods, without bill generation, balance tracking or payment input
    ' MONTHLY_IGNORE_BAL: monthly billing periods, without bill generation, balance tracking or payment input
    Public Enum BillingMode
        MONTHLY = 0
        ADHOC = 1
        MONTHLY_IGNORE_BAL = 2
    End Enum

    Public Enum SiteXactionGroup
        ROYALTY = 1
        INSURANCE = 10
        TAX = 20
        OTHER = 30
    End Enum
    Public Const QRY_RETAINSTATE As String = "retain"
    Public Const QRY_PREV_PG As String = "prevpg"
    Public Const QRY_SAVED As String = "saved"

    ' Attribute values
    ' GUI preferences or current states
    Public Const SESS_STUDENR As String = "CurEnr"


    ' Public Constants
    Public Const NULL_DATESTR As String = "1990-01-01"
    Public Shared NULL_DATE As Date = New Date(1990, 1, 1)
    Public Const NULL_GRADE As Integer = -99
    Public Const CUR_SCHOOL_END_YR As Integer = 2018
    Public Const MIN_GRAD_YR As Integer = 2000
    Public Const MAX_GRAD_YR As Integer = 2020
    Public Const DTRANGE_ALL_YEAR As Integer = 0
    Public Const DTRANGE_Q1_Q2 As Integer = -1
    Public Const DTRANGE_Q3_Q4 As Integer = -3


    ' Public constants for reference by child pages
    Public Const ENROLLED_ALL As Integer = -1
    Public Const ENROLLED_FALSE As Integer = 0
    Public Const ENROLLED_TRUE As Integer = 1
    Public Const SCOPE_TUTACTIVE As Integer = 0

    ' Application-wide
    Public Const MYAD_APPLICATION As String = "MyAdvantage"
    Public Const MYAD_CONNSTR As String = "myadConnectionString"
    Public Const MYAD_SEC_CONNSTR As String = "myadSecurityConnectionString"
    Public Const ATO_SEC_CONNSTR As String = "atoSecurityConnectionString"
    Public Const ATS_CONNSTR As String = "atsConnectionString"
    Public Const MEMPROV_MYAD As String = "MyAdSQLProvider"
    Public Const MEMPROV_ATO As String = "atoSQLProvider"
    Public Const ROLEPROV_MYAD As String = "MyAdSQLProvider"
    Public Const ROLEPROV_ATO As String = "atoSQLProvider"

    Public Const MYAD_TAG_URL_LOGIN As String = "loginurl"
    'Public Const MYAD_FILEPATH_LOGIN As String = "Login.aspx"
    'Public Const MYAD_FILEPATH_HOME As String = "Home.aspx"
    'Public Const MYAD_FILEPATH_CHPWD As String = "Password.aspx"
    Public Const URL_ATS As String = "http://www.at-ny.com"  '"http://capra"
    'Public Const URL_MYAD As String = "https://myad.at-ny.com"
    Public Const PGHDR As String = "pghdr"
    Public Const PGDVDR As String = "pgdvdr"
    Public Const TXT_ROW_SPACING As Integer = 4
    Public Const COL_PRIMSCHOOLS As String = "colPrim"
    Public Const COL_HIGHSCHOOLS As String = "colHS"
    Public Const COL_UNIVERSITIES As String = "colUniv"
    'Public Const COL_ROYALTYCATS As String = "colRoyCat"
    Public Const COL_INITIALS As String = "colInitials"
    Public Const COL_SESSIONVARS As String = "colSessionVars"

    ' Roles
    Public Const ROLE_ADM As String = "Admin"   ' Can see all test sched, result and cost info. Can sched tests
    Public Const ROLE_BILLADM As String = "BillAdmin"   ' Billing Admin -- sees all customer/acct info, sees comp rates
    Public Const ROLE_COMPADM As String = "CompAdmin"   ' Compensation admin -- sees all earnings/payroll info
    Public Const ROLE_CORPADM As String = "CorpAdmin"  ' View and edit posted revenue data for locations
    Public Const ROLE_CORPEXEC As String = "CorpExec" ' View company-wide reports
    Public Const ROLE_DIR As String = "Director"
    Public Const ROLE_EXEC As String = "Exec"
    Public Const ROLE_LOCALREVADM As String = "LocalRevAdmin"  ' Non-director who can see local site's revenue data
    Public Const ROLE_PCADM As String = "PCAdmin"   ' PC Admin - program consultant - sees customer info and comp rates
    Public Const ROLE_PROCTOR As String = "Proctor" ' Can see test sched and result info and can sched. Can't see costs.
    Public Const ROLE_TUT As String = "Tutor"
    Public Const ROLE_SYSADM As String = "SysAdmin"
    Public Const ROLE_TESTSYSADM As String = "TestSysAdmin"
    Public Const ROLE_TJOBADM As String = "TutorJobAdmin"
    Public Const ROLE_PJOBADM As String = "ProctorJobAdmin"
    Public Const ROLE_AJOBADM As String = "AdminJobAdmin"
    ' ATO-Specific Roles
    Public Const ROLE_OTUSER As String = "OTUser"
    ' Location Roles
    Public Const ROLE_AT_BOS As String = "AT-BOS"
    Public Const ROLE_AT_CHI As String = "AT-CHI"
    Public Const ROLE_AT_DC As String = "AT-DC"
    Public Const ROLE_AT_HST As String = "AT-HST"
    Public Const ROLE_AT_LA As String = "AT-LA"
    Public Const ROLE_AT_LI As String = "AT-LI"
    Public Const ROLE_AT_NNJ As String = "AT-NNJ"
    Public Const ROLE_AT_NYC As String = "AT-NYC"
    Public Const ROLE_AT_PAR As String = "AT-PAR"
    Public Const ROLE_AT_PHIL As String = "AT-PHIL"
    Public Const ROLE_AT_PTL As String = "AT-PTL"
    Public Const ROLE_AT_SFLA As String = "AT-SFLA"
    Public Const ROLE_AT_SV As String = "AT-SV"
    Public Const ROLE_AT_VT As String = "AT-VT"
    Public Const ROLE_AT_WEST As String = "AT-WEST"
    Public Const ROLE_AT_WPT As String = "AT-WPT"
    Public Const ROLE_AT_GUEST As String = "AT-GUEST"

    ' User Attributes
    Public Const USERNAME As String = "username"

    ' Student attributes
    Public Const STUDID As String = "StudentID"
    Public Const STUDENR As String = "Enrolled"
    Public Const STUDSCO As String = "Scope"
    Public Const STUDNAME As String = "StudentName"
    Public Const QRY_STUDID As String = "sid"
    Public Const QRY_STUDENR As String = "senr"
    Public Const QRY_STUDSCO As String = "sco"
    Public Const QRY_STUDNAME As String = "sn"
    Public Const QRY_STUDOPT As String = "sidopt"

    'Tutor Attributes
    Public Const TUTID As String = "TutorID"
    Public Const TUTSTAT As String = "TutorStatus"
    Public Const TUTNAME As String = "TutorName"
    Public Const TUTISTUT As String = "TutorIsTutor"
    Public Const TUTCOST As String = "TutorCost"
    Public Const TUTCOMP As String = "TutorComp"
    Public Const TUTSETDATE As String = "TutorSetDate"
    Public Const TUTNEEDRATEREVU As String = "TutorNeedRR"
    Public Const TUT_INITIALS As String = "Initials"
    Public Const TUT_TYPE As String = "tuttype"
    Public Const JADM_TYPES As String = "jadmtypes"
    Public Const QRY_TUTID As String = "tid"
    Public Const QRY_TUTNAME As String = "tn"
    Public Const QRY_TUTSTAT As String = "tstat"

    ' General
    Public Const QRY_USERNAME As String = "un"
    Public Const QRY_QRYSTRING As String = "QueryString"
    Public Const QRY_MSG As String = "msg"
    Public Const MSG As String = "msg"
    Public Const ERRMSG As String = "errmsg"
    Public Const STATMSG_TAG As String = "StatMsg"
    Public Const STATERR_TAG As String = "StatErr"
    Public Const ADD_SUCCEEDED As String = "AddSucc"

    ' STP-Related
    Public Const QRY_STPID As String = "stpid"
    Public Const STPID As String = "stpid"
    Public Const STPPROGID As String = "stpprogid"
    Public Const STPCOST As String = "stpcost"
    Public Const STPCOMP As String = "stpcomp"
    Public Const STPDEFLEN As String = "stpdeflen"
    Public Const STPSETDT As String = "stpsetdt"
    Public Const STPASS1ID As String = "stpass1id"
    Public Const STPASS1DT As String = "stpass1dt"
    Public Const STPASS2ID As String = "stpass2id"
    Public Const STPASS2DT As String = "stpass2dt"
    Public Const STPONHOLD As String = "stponhold"
    Public Const STPSTARTDT As String = "stpstartdt"

    ' TH-related
    Public Const QRY_THID As String = "thid"
    Public Const THID As String = "thid"
    Public Const TTID As String = "ttid"

    ' Site-Related
    Public Const QRY_SITEID As String = "stid"
    Public Const SITEID As String = "SiteID"    'DON'T CHANGE. This session param is hard-coded in many places
    Public Const INITIAL_SITEID As String = "InitSiteID"
    Public Const QRY_SITE_NAME As String = "stn"
    Public Const SITE_NAME As String = "SiteName"
    Public Const QRY_SREVID As String = "srid"
    Public Const SREVID As String = "srevid"
    Public Const SITE_SESS_LEN As String = "SiteSessionLength"
    Public Const SITE_FIN_RATE As String = "SiteFinanceRate"
    Public Const SITE_SET_DATE As String = "SiteSetDate"
    Public Const SITE_TUT_COST_ACCESS As String = "SiteTutCostAccess"
    Public Const SITE_TUT_ESSAY_MOD As String = "SiteEssayMod"
    Public Const SITE_RPT_CUTOVER_DAY As String = "SiteRptCutoverDay"

    ' SiteRev
    'Public Const SITEREV_STATE As String = "SiteRevState"

    ' Test-Related
    Public Const QRY_TSESS As String = "tsid"
    Public Const QRY_DTSTART As String = "rstrt"
    Public Const QRY_DTEND As String = "rend"
    Public Const QRY_DTIGNORE As String = "rign"
    Public Const QRY_CATID As String = "catid"
    Public Const QRY_TESTID As String = "tstid"
    Public Const QRY_TUTTAKER As String = "ttak"
    Public Const QRY_ONLINE_TEST As String = "ot"

    'Public Const RANGE_START As String = "rngst"
    'Public Const RANGE_END As String = "rngend"
    'Public Const RANGE_IGNORE As String = "rngign"

    ' Class-Related
    Public Const QRY_CLASSID As String = "cid"
    Public Const CLSID As String = "clsid"

    ' ATO-Related
    Public Const QDIAG_PATH As String = "~/images/diagrams/"
    Public Const QRY_OTQAID As String = "otqaid"
    Public Const QRY_TOKEN As String = "tok"

    ' When the finadmin starts processing payroll, COMP_YEAR/MONTH gets incremented so that new hours 
    '  entered are stamped with the next month, but COMP_REVIEW is set to 1, which will cause the finadmin
    '  comp-related pages to display the prev month by default.
    Public Const COMP_YEAR As String = "CPYear"
    Public Const COMP_MONTH As String = "CPMonth"
    Public Const COMP_MONTH_NAME As String = "CPMonthName"
    Public Const COMP_REVIEW As String = "CompReview"

    ' BILLING RELATED
    Public Const BILLING_MODE As String = "BillingMode"
    Public Const BILL_GENERATION As String = "BillGeneration"
    Public Const HRS_REVU_EXPORT As String = "HrsReviewExport"
    Public Const POS_COLLECTION As String = "PosCollection"

    ' If, while still processing billing for BP_YEAR/MONTH, the finadmin would like tutors to be able to 
    '   conveniently enter hours for the next month, HR_ENTRY_AHEAD is set to 1, which will cause the 
    '   hours entry and review pages to display the next month for all users who are not finadmins.
    Public Const BP_YEAR As String = "BPYear"       'Don't change. Used in Home.aspx
    Public Const BP_MONTH As String = "BPMonth"     'Don't change. Used in Home.aspx
    Public Const BP_MONTH_NAME As String = "BPMonthName"
    Public Const HRS_ENTRY_AHEAD As String = "HRSAhead"

    'Public Const HRS_ENTRY_YEAR As String = "HRSYear"
    'Public Const HRS_ENTRY_MONTH As String = "HRSMonth"
    'Public Const HRS_ENTRY_MONTH_NAME As String = "HRSMonthName"

    ' Ad Hoc Billing Period (BPAH)
    Public Const BPAH_PREV_END As String = "BPAHPrevEnd"
    Public Const BPAH_CUR_END As String = "BPAHCurEnd"

    ' Site Revenue related
    Public Const REV_YEAR As String = "REVYr"
    Public Const REV_MONTH As String = "REVMo"
    Public Const REV_MONTH_NAME As String = "REVMoName"

    ' Date Related
    Public Const QRY_YEAR As String = "yr"
    Public Const QRY_MONTH As String = "mo"
    Public Const QRY_ANNUAL As String = "ann"
    Public Const SEL_DATE As String = "seldate"

    ' Job Related
    Public Const QRY_JOBID As String = "jb"
    Public Const QRY_JTAID As String = "jtaid"
    Public Const JOBID As String = "jbid"
    Public Const JOBNAME As String = "jbname'"
    Public Const JTAID As String = "JTAID"

    ' SQL Fragments
    Public Const sqlfrag_CASE_MONTH As String = _
" WHEN 0 THEN 'NEW' WHEN 1 THEN 'JAN' " & _
" WHEN 2 THEN 'FEB' WHEN 3 THEN 'MAR' WHEN 4 THEN 'APR' WHEN 5 THEN 'MAY' WHEN 6 THEN 'JUN' WHEN 7 THEN 'JUL'" & _
" WHEN 8 THEN 'AUG' WHEN 9 THEN 'SEP' WHEN 10 THEN 'OCT' WHEN 11 THEN 'NOV' WHEN 12 THEN 'DEC' ELSE '???'" & _
" END + ' ' + "

    Public Const sqlfrag_CASE_COMP_PERIOD As String = _
"CASE CompMonth WHEN 0 THEN 'NEW' WHEN 1 THEN 'JAN' " & _
" WHEN 2 THEN 'FEB' WHEN 3 THEN 'MAR' WHEN 4 THEN 'APR' WHEN 5 THEN 'MAY' WHEN 6 THEN 'JUN' WHEN 7 THEN 'JUL'" & _
" WHEN 8 THEN 'AUG' WHEN 9 THEN 'SEP' WHEN 10 THEN 'OCT' WHEN 11 THEN 'NOV' WHEN 12 THEN 'DEC' ELSE '???'" & _
" END + ' ' + CAST(CompYear AS CHAR(4)) AS CompPeriod"

    Public Const sqlfrag_CASE_GRADE As String = _
"CASE CurrentGrade WHEN -1 THEN 'Unknown' WHEN 0 THEN 'Kindergarten' WHEN 1 THEN 'First' " & _
" WHEN 9 THEN 'HS Freshman' WHEN 10 THEN 'HS Sophomore' WHEN 11 THEN 'HS Junior' WHEN 12 THEN 'HS Senior' " & _
" WHEN 13 THEN 'Col Freshman' WHEN 14 THEN 'Col Soph' WHEN 15 THEN 'Col Junior' WHEN 16 THEN 'Col Senior' " & _
" WHEN 17 THEN 'Col Grad' ELSE CAST(CurrentGrade AS CHAR) END CurGrade"

    Public Const sqlfrag_EXCLUDE_TESTDATA_TUTS As String = _
" t.FullName<>'Cohan, Elizabeth' AND t.FullName<>'Lesem, Tina' " & _
" AND  NOT (t.LastName LIKE 'Schwartz%' and t.FirstName LIKE 'Jed%')  " 

    ' Shared between BillList.aspx.vb and rpt\BillList-Rpt.aspx.vb
    Public Const BLRPT_ITEM_NAME As Integer = 1
    Public Const BLRPT_ITEM_AMT As Integer = 5
    Public Const BLRPT_ITEM_DUE1 As Integer = 6
    Public Const BLRPT_ITEM_DUE2 As Integer = 7
    Public Const BLRPT_ITEM_DUE3 As Integer = 8
    Public Const BLRPT_ITEM_ADJ As Integer = 10
    Public Const BLRPT_ITEM_FIN As Integer = 11
    Public Const BLRPT_ITEM_BMETH As Integer = 13
    Public Const BLRPT_CELL_NAME As Integer = 0
    Public Const BLRPT_CELL_AUTO As Integer = 1
    Public Const BLRPT_CELL_AMT As Integer = 2
    Public Const BLRPT_CELL_DUE1 As Integer = 3
    Public Const BLRPT_CELL_DUE2 As Integer = 4
    Public Const BLRPT_CELL_DUE3 As Integer = 5
    Public Const BLRPT_CELL_ADJ As Integer = 8
    Public Const BLRPT_CELL_FIN As Integer = 9
    Public Const BLRPT_CELL_CON As Integer = 10
    Public Const BLRPT_CELL_EXP As Integer = 11
    Public Const BLRPT_CELL_GEN As Integer = 12

    ' Used by Utils, Hours-Add and Hours-Edit
    Public Const COLSUBINFO_TAG As String = "colSubInfo"

    ' Error-reporting (global.asx, MyAd-Error.aspx)
    Public Const ERR_URL As String = "errURL"
    Public Const ERR_MSG1 As String = "errMsg1"
    Public Const ERR_MSG2 As String = "errMsg2"




    Public Const PGSTAT_COL As String = "colPageState"
    Public Shared arrPGSTAT_VARS(,) As String = _
        {{STUDID, QRY_STUDID}, _
         {STUDENR, QRY_STUDENR}, _
         {STUDSCO, QRY_STUDSCO}, _
         {TUTID, QRY_TUTID}}


    ' Shared between AccountsReceivable-Summary.aspx.vb and rpt\AccountsReceivable-Summary-Rpt.aspx.vb
    Public Const ARSRPT_CELL_NAME As Integer = 0
    Public Const ARSRPT_ITEM_NAME As Integer = 1
    Public Const ARSRPT_CELL_AMT As Integer = 2
    Public Const ARSRPT_CELL_PAYMENT As Integer = 3
    Public Const ARSRPT_CELL_BALDUE As Integer = 4
    Public Const ARSRPT_CELL_BILLDATE As Integer = 5




End Class
