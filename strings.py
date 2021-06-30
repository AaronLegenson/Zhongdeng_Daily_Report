# -*- coding: utf-8 -*-

STRING_EMPTY = ""
STRING_SPACE = " "
STRING_PLAIN_ENTER = "\r\n"
STRING_HTML = "html"

STRING_FORMAT_3 = "{0}{1}{2}"

STRING_FORMAT_FULL_SECOND = "%Y-%m-%d %H:%M:%S"
STRING_FORMAT_FULL_SECOND_CLEAR = "%Y%m%d_%H%M%S"
STRING_FORMAT_FULL_MINUTE = "%Y-%m-%d %H:%M"
STRING_FORMAT_FULL_MICRO_SECOND = "%Y-%m-%d %H:%M:%S,%f"
STRING_FORMAT_NO_HYPHEN = "%Y%m%d"

STRING_COL_ORG_NAME = "查询企业名称"
STRING_COL_QUERY_TIME = "查询时间"
STRING_COL_KEYWORD_COUNT = "关键字数量"
STRING_COL_BUSINESS_ID = "机构id"
STRING_COL_BUSINESS_NAME = "机构名称"
STRING_COL_USER_ID = "用户id"
STRING_COL_USER_NAME = "用户名称"
STRING_COL_QUERY_COUNT = "搜索记录总次数"
STRING_COL_REGISTRATION_COUNT = "登记笔数"
STRING_COL_FILE_COUNT = "文件数"
STRING_COL_PAGE_COUNT = "页数"
STRING_COL_RELEASE_DATE = "开通时间"
STRING_COL_ACCOUNT_TYPE = "开通账户类别"
STRING_COL_ACCOUNT_COUNT = "开通账户数量"
STRING_COL_TEAM = "所属团队"
STRING_COL_TEAM_CONTACT = "团队对接人"
STRING_COL_RANK = "排名"
STRING_COL_WHETHER_FP = "是否开通发票验真账户"
STRING_COL_ROLES = "权限字段"
STRING_COL_QUERY_TIMES_IN_TEN_DAYS = "近十天查询次数"
STRING_COL_QUERY_TIMES_ALL = "总查询次数"
STRING_COL_QUERY_ORG_COUNT = "总查询企业数量"

STRING_SHEET_INSTRUCTIONS = "使用说明"
STRING_SHEET_BUSINESS = "开通机构信息"
STRING_SHEET_ACTIVITY = "近十天机构查询活跃度排名"
STRING_SHEET_DETAILS = "全量查询日志"
STRING_SHEET_USER = "开通用户信息"
STRING_SHEET_NEWLY_ADDED = "本期新增查询日志 [{0}至{1}]"
STRING_SHEET_NEWLY_ADDED_EMPTY = "本期新增查询日志"

STRING_EXCEL_RETURN_HOME = "返回主页"

STRING_CN_ENTIRE_COMPANY = "全公司"

# about OA
STRING_CHROME_OPTION_HEADLESS = "--headless"
STRING_FORMAT_7 = "{0}{1}{2}{3}{4}{5}{6}"
STRING_CHROMEDRIVER = "chromedriver"
STRING_LOGIN_USERNAME_CLASS = "lui_login_input_username"
STRING_LOGIN_PASSWORD_CLASS = "lui_login_input_password"
STRING_LOGIN_BUTTON_CLASS = "lui_login_button_div_l"
STRING_FORMAT_COOKIE = "{0}={1}; "
STRING_NAME = "name"
STRING_VALUE = "value"
STRING_COOKIE = "Cookie"
PARAMS_OA_USERNAME = "yunying_hook"
PARAMS_OA_PASSWORD = "Juliang123$%"
PARAMS_URL_BUILD_ADDRESS = \
    "https://oa.fusionfintrade.com/sys/common/dataxml.jsp?s_bean=sysZoneAddressTree&parent={0}&orgType=11&top="
PARAMS_URL_OA = "https://oa.fusionfintrade.com/login.jsp"
STRING_UTF_8 = "utf-8"
TERRIBLE_XML = ["\r", "\n", "\t", "\\r", "\\n", "\\t"]
TERRIBLE_DEFAULT = [" ", "\">", "&quot;"]
TRANS_DEFAULT = [["&lt;", "<"], ["&gt;", ">"]]
PARAMS_ALL_DEPARTMENTS = [
    "公司总经理",
    "金融机构合作首席代表刘志刚团队",
    "金融机构合作首席代表沈彦炜团队",
    "金融机构合作首席代表刘明团队",
    "金融机构合作首席代表张俊涛团队",
    "金融机构合作首席代表王晓光团队",
    "金融机构合作首席代表何伟强团队",
    "金融机构合作首席代表孟庆波团队",
    "医药行业事业部",
    "大数据管理部",
    "运营管理部",
    "技术开发部",
    "数字风控部",
    "计划财务部",
    "综合管理部",
    "集团董事长办公室",
    "增值服务业务部",
    "金融机构合作首席代表刘芊妤团队(筹)",
    "金融机构合作首席代表刘芊妤团队"
]
STRING_CN_DEPARTMENT = "部门"
STRING_CN_NAME = "姓名"


STRING_MAIL_TITLE = "SAM产品各首代团队客户使用日报"
STRING_MAIL_TEXT_ALIGN_RIGHT = " style=\"text-align: right;\""
STRING_MAIL_TEXT_HEAD = """
<html>
<head>
    <style>
    #table_normal table {
        width: 100%;
        margin: 15px 0;
        border: 0;
    }
    #table_normal th {
        background-color: #acd6ff;
        color:#000000
    }
    #table_normal,#table_normal th,#table_normal td {
        font-size: 1.0em;
        text-align: center;
        padding: 4px;
        border-collapse: collapse;
    }
    #table_normal th,#table_normal td {
        border: 1px solid #c4e1ff;
        border-width:1px 0 1px 0
    }
    #table_normal tr {
        border: 1px solid #c4e1ff;
    }
    #table_normal tr:nth-child(odd){
        background-color: #ecf5ff;
    }
    #table_normal tr:nth-child(even){
        background-color: #fdfdfd;
    }
    
    #table_highlight table {
        width: 100%;
        margin: 15px 0;
        border: 0;
    }
    #table_highlight th {
        background-color: #ff9797;
        color: #000000
    }
    #table_highlight,#table_highlight th,#table_highlight td {
        font-size: 1.0em;
        text-align: center;
        padding: 4px;
        border-collapse: collapse;
    }
    #table_highlight th,#table_highlight td {
        border: 1px solid #ffb5b5;
        border-width: 1px 0 1px 0
    }
    #table_highlight tr {
        border: 1px solid #ffb5b5;
    }
    #table_highlight tr:nth-child(odd){
        background-color: #ffecec;
    }
    #table_highlight tr:nth-child(even){
        background-color: #fdfdfd;
    }
    </style>
</head>
<body>
"""
STRING_MAIL_TEXT_TITLE = """
<h1>SAM产品各首代团队客户使用日报-{0}</h1>
"""

STRING_MAIL_TEXT_README = """
<h2>日报说明</h2>
<pre style="margin: 15px 0 0 0; padding: 20px; border: 0; border: 1px solid #c4e1ff; background: #ecf5ff; line-height: 1.4; font-family: Consolas; display: block; text-align: left; font-size: 1.0em; width: 1000px">
1. 欢迎查收《SAM产品各首代团队客户使用日报》
1. 本日报更新时间为每周一/三/五上午，日报内容截至更新日0时整
2. 如有疑问，请联系增值服务部：翁慧莹wenghuiying@fusionfintrade.com
</pre>
"""

STRING_MAIL_TEXT_TAIL = """
</body>
</html>
"""

STRING_MAIL_TEXT_PART_NORMAL = """
<h2>{0}</h2>
<pre style="margin: 15px 0 0 0; border: 0; line-height: 1.4; font-family: Consolas; display: block; text-align: center;">
{1}
</pre>
"""

STRING_MAIL_TEXT_PART_LOG = """
<h2>{0}</h2>
<pre style="margin: 15px 0 0 0; padding: 20px; border: 0; border: 1px solid #c4e1ff; background: #ecf5ff; line-height: 1.4; font-family: Consolas; display: block; text-align: left; font-size: 1.0em; width: 1000px">
{1}</pre>
"""

STRING_MAIL_TEXT_PART_NONE_BLUE = """
<h2>{0}</h2>
<pre style="margin: 15px 0 0 0; padding: 20px; border: 0; border: 1px solid #c4e1ff; background: #ecf5ff; line-height: 1.4; font-family: Consolas; display: block; text-align: center; font-size: 1.0em; width: 1000px">
{1}</pre>
"""

STRING_MAIL_TEXT_PART_NONE_RED = """
<h2>{0}</h2>
<pre style="margin: 15px 0 0 0; padding: 20px; border: 0; border: 1px solid #ffb5b5; background: #ffecec; line-height: 1.4; font-family: Consolas; display: block; text-align: center; font-size: 1.0em; width: 1000px">
{1}</pre>
"""

# STRING_MAIL_TEXT_PART_LOG = """
# <h2>{0}</h2>
# <pre style="margin: 15px 0 0 0; padding: 20px; border: 0; border: 1px dotted #785; background: #f5f5f5; line-height: 1.4; font-family: Consolas; display: block; text-align: left; font-size: 1.0em; width: 1000px">
# {1}</pre>
# """





